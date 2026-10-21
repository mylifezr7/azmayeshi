'use strict';
// Folder scanning + the local "ledger" that remembers what was already published,
// so a scheduled run never posts the same file twice.
const fs = require('node:fs');
const path = require('node:path');

const VIDEO_EXT = new Set(['.mp4', '.mov', '.m4v', '.webm']);
const collator = new Intl.Collator('en', { numeric: true, sensitivity: 'base' });

function listVideoFiles(mediaDir, { excludeDirs = [], include = VIDEO_EXT } = {}) {
  const root = path.resolve(mediaDir);
  const excluded = new Set([root, ...excludeDirs.map((dir) => path.resolve(dir))]);
  const found = [];

  function walk(dir) {
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch (error) {
      if (error.code === 'ENOENT') return;
      throw error;
    }
    for (const entry of entries) {
      if (entry.name.startsWith('.')) continue;
      const absolute = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        if (excluded.has(path.resolve(absolute))) continue;
        walk(absolute);
        continue;
      }
      if (!entry.isFile()) continue;
      if (!include.has(path.extname(entry.name).toLowerCase())) continue;
      const stat = fs.statSync(absolute);
      found.push({
        absolute,
        relative: path.relative(root, absolute),
        name: entry.name,
        size: stat.size,
        mtimeMs: stat.mtimeMs,
      });
    }
  }

  walk(root);
  return found.sort((a, b) => collator.compare(a.relative, b.relative));
}

function emptyState() {
  return { version: 1, posted: {}, attempts: {}, dailyCounters: {} };
}

function loadState(file) {
  try {
    const parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
    return { ...emptyState(), ...parsed, posted: parsed.posted || {}, attempts: parsed.attempts || {} };
  } catch (error) {
    if (error.code === 'ENOENT') return emptyState();
    throw new Error(`فایل وضعیت (${file}) خراب است: ${error.message}`);
  }
}

function saveState(file, state) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(state, null, 2));
  fs.renameSync(tmp, file);
}

function isPosted(state, file) {
  const record = state.posted[file.relative];
  return Boolean(record && record.mediaId);
}

function recordPosted(state, file, { mediaId, containerId, caption, permalink = null, at = new Date() }) {
  state.posted[file.relative] = {
    mediaId,
    containerId,
    caption,
    permalink,
    postedAt: at.toISOString(),
    size: file.size,
    mtimeMs: file.mtimeMs,
  };
  const day = at.toISOString().slice(0, 10);
  state.dailyCounters[day] = (state.dailyCounters[day] || 0) + 1;
  delete state.attempts[file.relative];
  return state;
}

function recordFailure(state, file, message, at = new Date()) {
  const previous = state.attempts[file.relative] || { count: 0 };
  state.attempts[file.relative] = {
    count: previous.count + 1,
    lastError: String(message).slice(0, 500),
    lastAt: at.toISOString(),
  };
  return state;
}

// Posts in the rolling last 24h, used to respect Instagram's publishing quota.
function postsInLast24h(state, now = new Date()) {
  const cutoff = now.getTime() - 24 * 60 * 60 * 1000;
  return Object.values(state.posted).filter((record) => {
    const at = Date.parse(record.postedAt || '');
    return Number.isFinite(at) && at >= cutoff;
  }).length;
}

function pendingFiles(files, state, { force = false } = {}) {
  return files.filter((file) => force || !isPosted(state, file));
}

function archiveFile(file, archiveDir) {
  fs.mkdirSync(archiveDir, { recursive: true });
  let target = path.join(archiveDir, file.name);
  if (fs.existsSync(target)) {
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    target = path.join(archiveDir, `${path.basename(file.name, path.extname(file.name))}-${stamp}${path.extname(file.name)}`);
  }
  try {
    fs.renameSync(file.absolute, target);
    return target;
  } catch (error) {
    if (error.code !== 'EXDEV') throw error;
    fs.copyFileSync(file.absolute, target);
    fs.unlinkSync(file.absolute);
    return target;
  }
}

module.exports = {
  listVideoFiles,
  loadState,
  saveState,
  isPosted,
  recordPosted,
  recordFailure,
  postsInLast24h,
  pendingFiles,
  archiveFile,
  emptyState,
  VIDEO_EXT,
  collator,
};
