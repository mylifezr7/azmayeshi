'use strict';
// One "run" = find new videos in the folder, publish them as Reels, remember them.
const fs = require('node:fs');
const path = require('node:path');
const { createMediaServer } = require('./media-server');
const { loadCaptionSources, resolveCaption } = require('./captions');
const { listVideoFiles, pendingFiles, recordPosted, recordFailure, postsInLast24h, saveState, archiveFile } = require('./library');
const { probeVideo, checkReelConstraints } = require('./probe');
const { sleep } = require('./graph');

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function acquireLock(file) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (fs.existsSync(file)) {
    try {
      const info = JSON.parse(fs.readFileSync(file, 'utf8'));
      if (info.pid && info.pid !== process.pid) {
        process.kill(info.pid, 0); // throws if the process is gone
        throw new Error(
          `یک اجرای دیگر از ربات در حال کار است (PID ${info.pid}). اگر مطمئنی اجرای قبلی بسته شده، فایل ${file} را پاک کن.`
        );
      }
    } catch (error) {
      if (error.code === 'EPERM' || error.code === 'ESRCH') {
        /* stale lock, safe to replace */
      } else if (!error.code) {
        throw error;
      }
    }
  }
  fs.writeFileSync(file, JSON.stringify({ pid: process.pid, startedAt: new Date().toISOString() }));
  let released = false;
  return () => {
    if (released) return;
    released = true;
    try {
      fs.unlinkSync(file);
    } catch {
      /* already gone */
    }
  };
}

async function publishOnce({
  config,
  client,
  state,
  logger = console.log,
  limit = config.maxPostsPerRun,
  dryRun = config.dryRun,
  force = false,
  archive = config.archiveAfterPost,
  onlyFile = null,
  signal = null,
  intervalMs = config.postIntervalSec * 1000,
  probe = true,
} = {}) {
  const summary = { published: [], failed: [], skipped: [], warnings: [] };

  const files = listVideoFiles(config.mediaDir, { excludeDirs: [config.archiveDir] });
  let candidates = pendingFiles(files, state, { force });
  if (onlyFile) {
    const wanted = path.basename(onlyFile);
    candidates = candidates.filter((file) => file.name === wanted || file.relative === onlyFile);
  }
  summary.skipped = files.filter((file) => !candidates.includes(file)).map((file) => file.relative);

  if (!candidates.length) {
    logger('ویدیوی جدیدی برای انتشار پیدا نشد.');
    return summary;
  }

  const alreadyToday = postsInLast24h(state);
  if (alreadyToday >= config.dailyLimit) {
    logger(`سقف انتشار ۲۴ ساعته پر شده است (${alreadyToday}/${config.dailyLimit}).`);
    return summary;
  }

  const sources = loadCaptionSources(config.mediaDir);
  let mediaServer = null;

  try {
    if (!dryRun) {
      mediaServer = createMediaServer({
        mediaDir: config.mediaDir,
        secret: config.mediaSecret,
        port: config.mediaPort,
        host: config.mediaHost,
        publicBaseUrl: config.publicBaseUrl,
        logger,
      });
      await mediaServer.start();
    }

    const batch = candidates.slice(0, Math.max(0, limit));
    for (let index = 0; index < batch.length; index += 1) {
      if (signal?.aborted) break;
      const file = batch[index];
      const { caption, cover, thumbOffset } = resolveCaption({
        relativePath: file.relative,
        mediaDir: config.mediaDir,
        config,
        sources,
        index: index + 1,
      });

      logger(`\n▶ ${file.relative} (${(file.size / 1024 ** 2).toFixed(1)} مگابایت)`);
      logger(`  کپشن: ${caption.split('\n')[0].slice(0, 80)}${caption.length > 80 ? '…' : ''}`);

      if (probe) {
        const info = await probeVideo(file.absolute, { command: config.ffprobe });
        const warnings = checkReelConstraints(file, info);
        for (const warning of warnings) logger(`  ⚠ ${warning}`);
        if (warnings.length) summary.warnings.push({ file: file.relative, warnings });
      }

      if (dryRun) {
        logger('  (dry-run: فقط بررسی شد، چیزی منتشر نشد)');
        summary.published.push({ file: file.relative, dryRun: true, caption });
        continue;
      }

      let containerId = null;
      try {
        const videoUrl = mediaServer.urlFor(file.relative);
        containerId = await client.createReelContainer({
          videoUrl,
          caption,
          shareToFeed: config.shareToFeed,
          coverUrl: cover || undefined,
          thumbOffset: thumbOffset ?? undefined,
        });
        logger(`  container ساخته شد: ${containerId}`);

        await client.waitForContainer(containerId, {
          intervalMs: 10_000,
          timeoutMs: 20 * 60_000,
          onStatus: ({ code, attempt }) => {
            if (attempt === 1 || attempt % 3 === 0) logger(`  وضعیت: ${code}`);
          },
        });

        const mediaId = await client.publish(containerId);
        recordPosted(state, file, { mediaId, containerId, caption });
        saveState(config.stateFile, state);
        summary.published.push({ file: file.relative, mediaId, containerId, caption });
        logger(`  ✅ منتشر شد (media id: ${mediaId})`);

        if (archive) {
          const target = archiveFile(file, config.archiveDir);
          logger(`  📦 به ${path.relative(config.mediaDir, target)} منتقل شد`);
        }
      } catch (error) {
        recordFailure(state, file, error.message);
        try {
          saveState(config.stateFile, state);
        } catch {
          /* state is best-effort on failure */
        }
        summary.failed.push({ file: file.relative, error: error.message, containerId });
        logger(`  ❌ شکست خورد: ${error.message}`);
        if (/rate|limit|quota|ظرفیت|2\)/i.test(error.message)) {
          logger('  توقف دسته به‌خاطر محدودیت نرخ اینستاگرام.');
          break;
        }
      }

      const isLast = index === batch.length - 1;
      if (!isLast && intervalMs > 0) {
        logger(`  ⏳ ${Math.round(intervalMs / 1000)} ثانیه تا ویدیوی بعدی…`);
        await delay(intervalMs);
      }
    }
  } finally {
    await mediaServer?.close();
  }

  return summary;
}

module.exports = { publishOnce, acquireLock };
