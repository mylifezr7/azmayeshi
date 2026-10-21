'use strict';
// Loads settings from environment variables and optional .env files.
// Nothing here talks to the network; validation happens only when a command needs it.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const MODULE_ROOT = path.resolve(__dirname, '..');
const REPO_ROOT = path.resolve(MODULE_ROOT, '..');

function loadEnvFile(file) {
  if (!fs.existsSync(file)) return;
  const text = fs.readFileSync(file, 'utf8');
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const match = line.match(/^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!match) continue;
    let value = match[2].trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (process.env[match[1]] === undefined) process.env[match[1]] = value;
  }
}

// Repo root first, then the module's own .env wins for anything still unset.
loadEnvFile(path.join(REPO_ROOT, '.env'));
loadEnvFile(path.join(MODULE_ROOT, '.env'));

function toBool(value, fallback = false) {
  if (value === undefined || value === null || value === '') return fallback;
  return /^(1|true|yes|on|بله|فعال)$/i.test(String(value).trim());
}

function toInt(value, fallback) {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function readOrCreateSecret(file) {
  try {
    const existing = fs.readFileSync(file, 'utf8').trim();
    if (existing.length >= 16) return existing;
  } catch {
    /* first run */
  }
  const secret = crypto.randomBytes(16).toString('hex');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, secret, { mode: 0o600 });
  return secret;
}

function parseHashtags(value) {
  const list = String(value || '')
    .split(/[\s,]+/)
    .map((tag) => tag.trim())
    .filter(Boolean)
    .map((tag) => (tag.startsWith('#') ? tag : `#${tag}`));
  return [...new Set(list)];
}

function loadConfig(overrides = {}) {
  const env = process.env;
  const dataDir = path.resolve(MODULE_ROOT, env.IG_DATA_DIR || 'data');
  const mediaDir = path.resolve(MODULE_ROOT, overrides.mediaDir || env.IG_MEDIA_DIR || 'media');
  const archiveDir = path.resolve(mediaDir, env.IG_ARCHIVE_DIR || 'posted');
  const secret = env.IG_MEDIA_SECRET || readOrCreateSecret(path.join(dataDir, 'media-secret.txt'));

  return {
    moduleRoot: MODULE_ROOT,
    mediaDir,
    archiveDir,
    dataDir,
    stateFile: path.join(dataDir, 'state.json'),
    lockFile: path.join(dataDir, 'uploader.lock'),
    accessToken: env.IG_ACCESS_TOKEN || '',
    igUserId: env.IG_USER_ID || '',
    graphVersion: env.IG_GRAPH_VERSION || 'v23.0',
    graphBaseUrl: env.IG_GRAPH_BASE_URL || 'https://graph.facebook.com',
    tokenBaseUrl: env.IG_TOKEN_BASE_URL || 'https://graph.instagram.com',
    publicBaseUrl: (env.IG_PUBLIC_BASE_URL || '').replace(/\/+$/, ''),
    mediaSecret: secret,
    mediaPort: toInt(env.IG_MEDIA_PORT, 8787),
    mediaHost: env.IG_MEDIA_HOST || '0.0.0.0',
    defaultCaption: env.IG_DEFAULT_CAPTION || '{name}',
    defaultHashtags: parseHashtags(env.IG_DEFAULT_HASHTAGS),
    shareToFeed: toBool(env.IG_SHARE_TO_FEED, true),
    dailyLimit: toInt(env.IG_DAILY_LIMIT, 50),
    postIntervalSec: toInt(env.IG_POST_INTERVAL_SEC, 45),
    watchIntervalSec: toInt(env.IG_WATCH_INTERVAL_SEC, 300),
    maxPostsPerRun: toInt(env.IG_MAX_POSTS_PER_RUN, 5),
    archiveAfterPost: toBool(env.IG_ARCHIVE_AFTER_POST, false),
    dryRun: toBool(env.IG_DRY_RUN, false),
    ffprobe: env.IG_FFPROBE || 'ffprobe',
    ...overrides,
  };
}

// Human-readable problems, in Persian, so the CLI can explain what is missing.
function validateForPublishing(config) {
  const problems = [];
  if (!config.igUserId) problems.push('IG_USER_ID تنظیم نشده است (شناسهٔ اکانت بیزینسی اینستاگرام).');
  if (!config.accessToken) problems.push('IG_ACCESS_TOKEN تنظیم نشده است (توکن دسترسی Graph API).');
  if (!config.publicBaseUrl) {
    problems.push(
      'IG_PUBLIC_BASE_URL تنظیم نشده است. اینستاگرام ویدیو را فقط از یک URL عمومی HTTPS می‌خواند؛ ' +
        'اول تونل (مثل cloudflared tunnel --url http://localhost:8787) را بالا بیاور و آدرسش را اینجا بگذار.'
    );
  }
  if (!/^https:\/\//i.test(config.publicBaseUrl) && config.publicBaseUrl) {
    problems.push('IG_PUBLIC_BASE_URL باید با https:// شروع شود و از بیرون قابل دسترس باشد.');
  }
  return problems;
}

module.exports = { loadConfig, validateForPublishing, parseHashtags, toBool, toInt, MODULE_ROOT, REPO_ROOT };
