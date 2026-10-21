'use strict';
// Resolves the caption for each video.
// Priority: captions.json entry -> sidecar "<video>.txt" -> default template.
const fs = require('node:fs');
const path = require('node:path');

const CAPTION_LIMIT = 2200;
const HASHTAG_LIMIT = 30;

function readJsonSafe(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (error) {
    if (error.code === 'ENOENT') return null;
    throw new Error(`خواندن ${path.basename(file)} ناموفق بود: ${error.message}`);
  }
}

function loadCaptionSources(mediaDir) {
  const file = path.join(mediaDir, 'captions.json');
  const raw = readJsonSafe(file);
  const byFile = new Map();
  let fallback = null;

  if (Array.isArray(raw)) {
    for (const item of raw) {
      if (!item || typeof item !== 'object' || !item.file) continue;
      byFile.set(String(item.file), item);
    }
  } else if (raw && typeof raw === 'object') {
    for (const [key, value] of Object.entries(raw)) {
      if (key === '_default' || key === 'default') {
        fallback = typeof value === 'string' ? value : value?.caption ?? null;
        continue;
      }
      if (key.startsWith('_') || key.startsWith('//')) continue;
      byFile.set(key, typeof value === 'string' ? { caption: value } : value);
    }
  }

  return { byFile, fallback, file };
}

function prettifyName(fileName) {
  return path
    .basename(fileName, path.extname(fileName))
    .replace(/^\s*\d{1,4}\s*[-_.) ]+\s*/, '')
    .replace(/[_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function applyTemplate(template, { relativePath, fileName, index, date }) {
  return String(template || '')
    .replace(/\{name\}/g, prettifyName(fileName))
    .replace(/\{file\}/g, fileName)
    .replace(/\{path\}/g, relativePath)
    .replace(/\{index\}/g, String(index))
    .replace(/\{date\}/g, date);
}

function countHashtags(text) {
  return (String(text).match(/#[^\s#]+/g) || []).length;
}

function trimToLimit(text, limit = CAPTION_LIMIT) {
  if (text.length <= limit) return text;
  const cut = text.slice(0, limit);
  const boundary = cut.lastIndexOf(' ');
  return (boundary > limit * 0.6 ? cut.slice(0, boundary) : cut).trimEnd();
}

function appendHashtags(caption, hashtags) {
  if (!hashtags || !hashtags.length) return caption;
  const room = Math.max(0, HASHTAG_LIMIT - countHashtags(caption));
  const extras = hashtags.slice(0, room);
  if (!extras.length) return caption;
  return `${caption.trimEnd()}\n\n${extras.join(' ')}`;
}

function resolveCaption({ relativePath, mediaDir, config, sources, index = 1, date = new Date() }) {
  const fileName = path.basename(relativePath);
  const baseName = path.basename(relativePath, path.extname(relativePath));
  const isoDate = date.toISOString().slice(0, 10);

  const entry =
    sources?.byFile.get(relativePath) ||
    sources?.byFile.get(fileName) ||
    sources?.byFile.get(baseName) ||
    null;

  let caption = null;
  let cover = null;
  let thumbOffset = null;

  if (entry) {
    caption = typeof entry === 'string' ? entry : entry.caption ?? entry.text ?? null;
    cover = entry.cover ?? entry.cover_url ?? null;
    thumbOffset = entry.thumb_offset ?? entry.thumbOffset ?? null;
  }

  if (!caption) {
    const sidecar = ['.txt', '.caption.txt', '.md'].map((ext) => path.join(mediaDir, baseName + ext)).find((file) => fs.existsSync(file));
    if (sidecar) caption = fs.readFileSync(sidecar, 'utf8');
  }

  if (!caption && sources?.fallback) caption = sources.fallback;
  if (!caption) caption = config.defaultCaption || '{name}';

  caption = applyTemplate(caption, { relativePath, fileName, index, date: isoDate });
  if (countHashtags(caption) === 0) caption = appendHashtags(caption, config.defaultHashtags);
  caption = trimToLimit(caption.replace(/\r\n/g, '\n').trim());

  return { caption, cover, thumbOffset };
}

module.exports = {
  loadCaptionSources,
  resolveCaption,
  applyTemplate,
  appendHashtags,
  trimToLimit,
  countHashtags,
  prettifyName,
  CAPTION_LIMIT,
  HASHTAG_LIMIT,
};
