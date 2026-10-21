#!/usr/bin/env node
'use strict';
// Instagram video (Reels) uploader — command line entry point.
// Usage: node instagram-uploader/cli.js <command> [options]
const fs = require('node:fs');
const path = require('node:path');
const { loadConfig, validateForPublishing } = require('./src/config');
const { createGraphClient, identifyAccount, refreshLongLivedToken } = require('./src/graph');
const { loadCaptionSources, resolveCaption } = require('./src/captions');
const { listVideoFiles, loadState, saveState, postsInLast24h, isPosted, emptyState } = require('./src/library');
const { probeVideo, checkReelConstraints } = require('./src/probe');
const { createMediaServer } = require('./src/media-server');
const { publishOnce, acquireLock } = require('./src/publisher');

const HELP = `
ربات انتشار ویدیو در پیج اینستاگرام (API رسمی Graph / Reels)

دستورها:
  whoami           توکن را بده، خودش اکانت و شناسه را پیدا و (با --save) ذخیره می‌کند
  check            بررسی تنظیمات، توکن و دسترسی عمومی به فایل ویدیو
  list             نمایش ویدیوهای پوشه، وضعیت انتشار و کپشن هرکدام
  post             انتشار ویدیوهای جدید (یک‌بار)
  watch            رصد مداوم پوشه و انتشار خودکار ویدیوهای تازه
  serve            فقط اجرای سرور فایل ویدیو (برای تست تونل)
  status           خلاصهٔ آمار، سهمیهٔ اینستاگرام و آخرین خطاها
  refresh-token    تمدید توکن ۶۰روزهٔ اینستاگرام (--save برای ذخیره در .env)
  reset            پاک‌کردن حافظهٔ انتشار (برای ارسال دوباره)

گزینه‌ها:
  --limit N        حداکثر تعداد انتشار در این اجرا
  --interval S     فاصلهٔ ثانیه‌ای بین دو انتشار (پیش‌فرض ۴۵)
  --file NAME      فقط همین فایل را منتشر کن
  --dry-run        فقط شبیه‌سازی؛ چیزی به اینستاگرام نمی‌رود
  --force          ویدیوهای قبلاً منتشرشده را هم دوباره بفرست
  --archive        بعد از انتشار موفق، فایل را به پوشهٔ posted/ منتقل کن
  --quiet          لاگ کمتر
  --save           ذخیرهٔ نتیجه در .env (whoami و refresh-token)
  --token TOKEN    توکن را مستقیم بده (whoami) به‌جای خواندن از .env

نمونه:
  node instagram-uploader/cli.js list
  node instagram-uploader/cli.js check
  node instagram-uploader/cli.js post --limit 1 --dry-run
  node instagram-uploader/cli.js watch --interval 600 --archive

تنظیمات از فایل instagram-uploader/.env خوانده می‌شود (نمونه: .env.example).
`;

function parseArgs(argv) {
  const options = { _: [] };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith('--')) {
      options._.push(arg);
      continue;
    }
    const [rawKey, inlineValue] = arg.slice(2).split('=');
    const key = rawKey.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
    if (inlineValue !== undefined) {
      options[key] = inlineValue;
    } else if (argv[i + 1] && !argv[i + 1].startsWith('--')) {
      options[key] = argv[++i];
    } else {
      options[key] = true;
    }
  }
  return options;
}

function preview(text, length = 90) {
  const single = String(text).replace(/\s+/g, ' ').trim();
  return single.length > length ? `${single.slice(0, length)}…` : single;
}

function makeClient(config) {
  return createGraphClient({
    accessToken: config.accessToken,
    igUserId: config.igUserId,
    version: config.graphVersion,
    baseUrl: config.graphBaseUrl,
  });
}

async function commandCheck(config, options) {
  const problems = validateForPublishing(config);
  if (problems.length) {
    console.log('⛔ تنظیمات ناقص است:\n' + problems.map((p) => `  - ${p}`).join('\n'));
    console.log(`\nمسیر رسانه: ${config.mediaDir}`);
    return 1;
  }

  const token = config.accessToken;
  const usesInstagramLogin = /graph\.instagram\.com/i.test(config.graphBaseUrl);
  const tokenProblems = [];
  if (/^IGAA/i.test(token)) {
    tokenProblems.push(
      'این توکن از «API Setup» قدیمی (Instagram Basic Display) است که از دسامبر ۲۰۲۴ مرده و هیچ پستی نمی‌پذیرد. ' +
        'باید از Instagram → API setup with Instagram login ساخته شود.'
    );
  } else if (/^IGQ/i.test(token) && !usesInstagramLogin) {
    tokenProblems.push('توکن از نوع Instagram Login است ولی IG_GRAPH_BASE_URL روی graph.facebook.com است؛ آن را https://graph.instagram.com بگذار.');
  } else if (/^EAA/i.test(token) && usesInstagramLogin) {
    tokenProblems.push('توکن از نوع Facebook است ولی IG_GRAPH_BASE_URL روی graph.instagram.com است؛ آن را https://graph.facebook.com بگذار.');
  }
  if (tokenProblems.length) {
    console.log('⛔ مشکل توکن:\n' + tokenProblems.map((p) => `  - ${p}`).join('\n'));
    return 1;
  }

  console.log('🔑 بررسی توکن و نوع اکانت…');
  const client = makeClient(config);
  const account = await client.getAccount();
  console.log(`  اکانت: @${account.username || '?'} (${account.name || 'بدون نام'})`);
  console.log(`  نوع: ${account.account_type || 'نامشخص'} — تعداد پست: ${account.media_count ?? '?'}`);
  let warning = false;
  if (!/BUSINESS|CREATOR|MEDIA_CREATOR/i.test(account.account_type || '')) {
    console.log('  ⚠ برای انتشار از API باید اکانت بیزینسی/کریتور و متصل به یک صفحهٔ فیسبوک باشد.');
    warning = true;
  }
  if (client.lastUsage) console.log(`  مصرف سهمیهٔ اپ: ${JSON.stringify(client.lastUsage)}`);
  try {
    const limit = await client.getPublishingLimit();
    if (limit.quotaTotal !== null) {
      console.log(`  سهمیهٔ انتشار ۲۴ ساعته: ${limit.quotaUsage ?? '?'} از ${limit.quotaTotal}`);
    }
  } catch (error) {
    console.log(`  (خواندن سهمیهٔ انتشار ممکن نشد: ${error.message})`);
  }

  const files = listVideoFiles(config.mediaDir, { excludeDirs: [config.archiveDir] });
  console.log(`📁 ${files.length} فایل ویدیویی در ${config.mediaDir}`);
  if (!files.length) {
    console.log('  (برای تست دسترسی عمومی حداقل یک ویدیو داخل پوشه بگذار.)');
    return warning ? 1 : 0;
  }

  const sample = files[0];
  const server = createMediaServer({
    mediaDir: config.mediaDir,
    secret: config.mediaSecret,
    port: config.mediaPort,
    host: config.mediaHost,
    publicBaseUrl: config.publicBaseUrl,
  });
  await server.start();
  try {
    const url = server.urlFor(sample.relative);
    console.log(`🌍 بررسی دسترسی عمومی به نمونه: ${url}`);
    try {
      const response = await fetch(url, { method: 'HEAD', signal: AbortSignal.timeout(20_000) });
      console.log(`  پاسخ: HTTP ${response.status} — ${response.headers.get('content-type') || 'بدون content-type'}`);
      if (!response.ok) {
        console.log('  ⛔ فایل از بیرون قابل خواندن نیست؛ تونل یا IG_PUBLIC_BASE_URL را بررسی کن.');
        warning = true;
      }
    } catch (error) {
      console.log(`  ⛔ از این دستگاه نمی‌شود به ${url} وصل شد: ${error.message}`);
      console.log('  (اگر تونل از بیرون کار می‌کند ولی از داخل نه، این خط را جدی نگیر.)');
      warning = true;
    }
  } finally {
    await server.close();
  }
  return warning ? 1 : 0;
}

function commandList(config, options) {
  const state = loadState(config.stateFile);
  const sources = loadCaptionSources(config.mediaDir);
  const files = listVideoFiles(config.mediaDir, { excludeDirs: [config.archiveDir] });
  if (!files.length) {
    console.log(`پوشهٔ ${config.mediaDir} خالی است.`);
    return 0;
  }
  console.log(`📁 ${files.length} ویدیو در ${config.mediaDir}\n`);
  files.forEach((file, index) => {
    const posted = isPosted(state, file);
    const { caption } = resolveCaption({
      relativePath: file.relative,
      mediaDir: config.mediaDir,
      config,
      sources,
      index: index + 1,
    });
    console.log(`${posted ? '✅' : '⏳'} ${file.relative}  (${(file.size / 1024 ** 2).toFixed(1)}MB)`);
    console.log(`   کپشن: ${preview(caption)}`);
    if (posted) console.log(`   منتشرشده: ${state.posted[file.relative].postedAt} — media ${state.posted[file.relative].mediaId}`);
    const attempt = state.attempts[file.relative];
    if (attempt) console.log(`   آخرین خطا (${attempt.count} بار): ${preview(attempt.lastError, 120)}`);
    if (options.verbose) {
      // probe is async; in verbose we only show file dates to stay synchronous
      console.log(`   mtime: ${new Date(file.mtimeMs).toISOString()}`);
    }
  });
  console.log(`\nامروز (۲۴ ساعت گذشته): ${postsInLast24h(state)} از سقف ${config.dailyLimit}`);
  return 0;
}

async function commandStatus(config) {
  const state = loadState(config.stateFile);
  const posted = Object.entries(state.posted);
  console.log(`تعداد کل انتشار ثبت‌شده: ${posted.length}`);
  console.log(`۲۴ ساعت گذشته (از حافظهٔ محلی): ${postsInLast24h(state)} از سقف محلی ${config.dailyLimit}`);
  if (config.accessToken && config.igUserId) {
    try {
      const limit = await makeClient(config).getPublishingLimit();
      if (limit.quotaTotal !== null) {
        console.log(`سهمیهٔ رسمی اینستاگرام: ${limit.quotaUsage ?? '?'} از ${limit.quotaTotal} در ۲۴ ساعت گذشته`);
      }
    } catch (error) {
      console.log(`سهمیهٔ رسمی اینستاگرام خوانده نشد: ${error.message}`);
    }
  }
  if (posted.length) {
    console.log('\nآخرین انتشارها:');
    posted
      .sort((a, b) => String(b[1].postedAt).localeCompare(String(a[1].postedAt)))
      .slice(0, 5)
      .forEach(([file, record]) => {
        console.log(`  ${record.postedAt}  ${file}  →  media ${record.mediaId}${record.permalink ? `  ${record.permalink}` : ''}`);
      });
  }
  const failures = Object.entries(state.attempts);
  if (failures.length) {
    console.log('\nخطاهای قبلی:');
    failures.forEach(([file, attempt]) => {
      console.log(`  ${file} — ${attempt.count} بار — ${preview(attempt.lastError, 120)}`);
    });
  }
  console.log(`\nوضعیت: ${config.dryRun ? 'DRY-RUN' : 'انتشار واقعی'} — فاصلهٔ پیش‌فرض ${config.postIntervalSec}s`);
  return 0;
}

async function withLock(config, fn) {
  const release = acquireLock(config.lockFile);
  try {
    return await fn();
  } finally {
    release();
  }
}

async function commandPost(config, options) {
  const dryRun = Boolean(options.dryRun) || config.dryRun;
  const problems = validateForPublishing(config).filter((problem) => !(dryRun && /IG_PUBLIC_BASE_URL/.test(problem)));
  if (problems.length) {
    console.log('⛔ تنظیمات ناقص است:\n' + problems.map((p) => `  - ${p}`).join('\n'));
    return 1;
  }

  if (!fs.existsSync(config.mediaDir)) {
    console.log(`پوشهٔ رسانه وجود ندارد: ${config.mediaDir}`);
    return 1;
  }

  const state = loadState(config.stateFile);
  const limit = Number(options.limit || config.maxPostsPerRun);

  return withLock(config, async () => {
    const summary = await publishOnce({
      config,
      client: dryRun ? { createReelContainer: null } : makeClient(config),
      state,
      limit,
      dryRun,
      force: Boolean(options.force),
      archive: Boolean(options.archive) || config.archiveAfterPost,
      onlyFile: typeof options.file === 'string' ? options.file : null,
      intervalMs: Number(options.interval || config.postIntervalSec) * 1000,
      logger: options.quiet ? () => {} : console.log,
    });
    saveState(config.stateFile, state);

    console.log(
      `\nخلاصه: ✅ ${summary.published.length} منتشر شد، ⏭ ${summary.skipped.length} از قبل بود، ❌ ${summary.failed.length} خطا`
    );
    return summary.failed.length ? 1 : 0;
  });
}

async function commandWatch(config, options) {
  const dryRun = Boolean(options.dryRun) || config.dryRun;
  const problems = validateForPublishing(config).filter((problem) => !(dryRun && /IG_PUBLIC_BASE_URL/.test(problem)));
  if (problems.length) {
    console.log('⛔ تنظیمات ناقص است:\n' + problems.map((p) => `  - ${p}`).join('\n'));
    return 1;
  }

  const everySec = Math.max(30, Number(options.interval || config.watchIntervalSec));
  const limit = Number(options.limit || Math.min(config.maxPostsPerRun, 1));
  const state = loadState(config.stateFile);
  const client = dryRun ? {} : makeClient(config);
  let running = true;

  const stop = () => {
    running = false;
    console.log('\n⏹ در حال توقف… (بعد از پایان کار جاری بسته می‌شود)');
  };
  process.on('SIGINT', stop);
  process.on('SIGTERM', stop);

  const release = acquireLock(config.lockFile);
  console.log(`👀 رصد ${config.mediaDir} هر ${everySec} ثانیه. برای خروج Ctrl+C.`);
  try {
    while (running) {
      const summary = await publishOnce({
        config,
        client,
        state,
        limit,
        dryRun,
        force: false,
        archive: Boolean(options.archive) || config.archiveAfterPost,
        intervalMs: Number(options.interval || config.postIntervalSec) * 1000,
        logger: options.quiet ? () => {} : console.log,
      });
      saveState(config.stateFile, state);
      if (summary.published.length || summary.failed.length) {
        console.log(
          `— این دور: ✅ ${summary.published.length}، ❌ ${summary.failed.length} — ${new Date().toLocaleString('fa-IR')}`
        );
      } else if (!options.quiet) {
        console.log(`— چیزی برای انتشار نبود (${new Date().toLocaleTimeString('fa-IR')})`);
      }
      if (!running) break;
      for (let waited = 0; waited < everySec && running; waited += 5) {
        await new Promise((resolve) => setTimeout(resolve, 5000));
      }
    }
  } finally {
    release();
    process.removeListener('SIGINT', stop);
    process.removeListener('SIGTERM', stop);
  }
  return 0;
}

async function commandServe(config) {
  const server = createMediaServer({
    mediaDir: config.mediaDir,
    secret: config.mediaSecret,
    port: config.mediaPort,
    host: config.mediaHost,
    publicBaseUrl: config.publicBaseUrl,
  });
  const { port } = await server.start();
  const files = listVideoFiles(config.mediaDir, { excludeDirs: [config.archiveDir] });
  console.log(`سرور روی پورت ${port} آماده است. نمونهٔ URL:`);
  console.log(`  ${files[0] ? server.urlFor(files[0].relative) : `http://127.0.0.1:${port}/${config.mediaSecret}/نام-فایل.mp4`}`);
  console.log('برای انتشار عمومی: cloudflared tunnel --url http://localhost:' + port);
  console.log('Ctrl+C برای خروج.');
  await new Promise((resolve) => {
    process.on('SIGINT', resolve);
    process.on('SIGTERM', resolve);
  });
  await server.close();
  return 0;
}

// Writes KEY=value pairs at the END of .env: the loader lets the last
// occurrence win, so this also fixes an older, wrong line above it.
function saveEnvValues(values, file) {
  let text = '';
  try {
    text = fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '');
  } catch {
    /* first write */
  }
  const lines = Object.entries(values)
    .filter(([, value]) => value !== undefined && value !== null && value !== '')
    .map(([key, value]) => `${key}=${String(value).replace(/[\r\n]/g, '')}`);
  if (!lines.length) return file;
  const updated = `${text.trimEnd()}\n${lines.join('\n')}\n`.replace(/^\n/, '');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, updated, { mode: 0o600 });
  return file;
}

async function commandWhoami(config, options) {
  const token = typeof options.token === 'string' ? options.token.trim() : config.accessToken;
  if (!token) {
    console.log('توکن نداری. با --token بده یا اول IG_ACCESS_TOKEN را در .env بگذار.');
    return 1;
  }
  if (/^IGAA/i.test(token)) {
    console.log('⛔ این توکن از «API Setup» قدیمی است و از دسامبر ۲۰۲۴ کار نمی‌کند. توکن تازه از «API setup with Instagram login» بگیر.');
    return 1;
  }
  const preferred = config.graphBaseUrlExplicit
    ? config.graphBaseUrl
    : /^IGQ/i.test(token)
      ? 'https://graph.instagram.com'
      : 'https://graph.facebook.com';
  console.log(`🔍 در حال پیدا کردن اکانت (${preferred}) …`);
  let account;
  let usedBase = preferred;
  try {
    account = await identifyAccount({ accessToken: token, baseUrl: preferred, version: config.graphVersion });
  } catch (error) {
    const alternative = preferred.includes('graph.instagram.com') ? 'https://graph.facebook.com' : 'https://graph.instagram.com';
    try {
      account = await identifyAccount({ accessToken: token, baseUrl: alternative, version: config.graphVersion });
      usedBase = alternative;
    } catch {
      throw error;
    }
  }

  console.log(`✅ پیدا شد: @${account.username || 'نامشخص'}${account.accountType ? ` (${account.accountType})` : ''}`);
  console.log(`   شناسهٔ اکانت (IG_USER_ID): ${account.id}`);
  console.log(`   آدرس API درست: ${usedBase}`);

  const envFile = path.join(__dirname, '.env');
  if (options.save) {
    const values = { IG_ACCESS_TOKEN: token, IG_USER_ID: account.id, IG_GRAPH_BASE_URL: usedBase };
    if (/graph\.instagram\.com/.test(usedBase)) values.IG_GRAPH_BASE_URL = 'https://graph.instagram.com';
    saveEnvValues(values, envFile);
    console.log(`💾 ذخیره شد در ${envFile}`);
  } else {
    console.log('برای ذخیرهٔ خودکار همین مقادیر: دوباره با --save اجرا کن.');
  }
  return 0;
}

function saveTokenToEnvFile(token, file) {
  return saveEnvValues({ IG_ACCESS_TOKEN: token }, file);
}

async function commandRefreshToken(config, options) {
  if (!config.accessToken) {
    console.log('IG_ACCESS_TOKEN تنظیم نشده است؛ چیزی برای تمدید نیست.');
    return 1;
  }
  const baseUrl = config.tokenBaseUrl || 'https://graph.instagram.com';
  console.log(`⏳ تمدید توکن از ${baseUrl} …`);
  const result = await refreshLongLivedToken({ accessToken: config.accessToken, baseUrl });
  const days = result.expiresInSec ? Math.floor(result.expiresInSec / 86400) : null;
  console.log(`✅ توکن تازه${days ? ` (${days} روز اعتبار)` : ''}:`);
  console.log(result.accessToken);
  if (options.save) {
    const envFile = path.join(__dirname, '.env');
    saveTokenToEnvFile(result.accessToken, envFile);
    console.log(`ذخیره شد در ${envFile} — ربات را دوباره اجرا کن.`);
  } else {
    console.log('برای ذخیرهٔ خودکار: --save (فایل instagram-uploader/.env)');
  }
  return 0;
}

function commandReset(config, options) {
  if (!options.all && typeof options.file !== 'string') {
    console.log('برای پاک‌کردن کل حافظه: reset --all\nبرای یک فایل: reset --file NAME');
    return 1;
  }
  if (options.all) {
    saveState(config.stateFile, emptyState());
    console.log('حافظهٔ انتشار پاک شد.');
    return 0;
  }
  const state = loadState(config.stateFile);
  let removed = 0;
  for (const key of Object.keys(state.posted)) {
    if (key === options.file || path.basename(key) === path.basename(options.file)) {
      delete state.posted[key];
      removed += 1;
    }
  }
  for (const key of Object.keys(state.attempts)) {
    if (key === options.file || path.basename(key) === path.basename(options.file)) {
      delete state.attempts[key];
      removed += 1;
    }
  }
  saveState(config.stateFile, state);
  console.log(removed ? `${removed} رکورد پاک شد.` : 'رکوردی پیدا نشد.');
  return 0;
}

async function main() {
  const argv = process.argv.slice(2);
  const options = parseArgs(argv);
  const command = options._[0] || 'help';

  if (command === 'help' || options.help) {
    console.log(HELP);
    return 0;
  }
  if (command === 'gen-secret') {
    const config = loadConfig();
    console.log(config.mediaSecret);
    return 0;
  }

  const config = loadConfig();

  switch (command) {
    case 'check':
      return commandCheck(config, options);
    case 'list':
      return commandList(config, options);
    case 'status':
      return commandStatus(config);
    case 'post':
      return commandPost(config, options);
    case 'watch':
      return commandWatch(config, options);
    case 'serve':
      return commandServe(config);
    case 'whoami':
      return commandWhoami(config, options);
    case 'refresh-token':
      return commandRefreshToken(config, options);
    case 'reset':
      return commandReset(config, options);
    default:
      console.log(`دستور ناشناخته: ${command}\n${HELP}`);
      return 1;
  }
}

main()
  .then((code) => {
    process.exitCode = code || 0;
  })
  .catch((error) => {
    console.error(`\n❌ ${error.message}`);
    if (process.env.IG_DEBUG) console.error(error.stack);
    process.exitCode = 1;
  });

module.exports = { parseArgs };
