'use strict';
// Tiny static server that exposes the videos in MEDIA_DIR at unguessable URLs,
// because Instagram downloads the video from a public URL instead of accepting an upload.
// Put a tunnel (cloudflared / ngrok / your own reverse proxy) in front of it.
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const VIDEO_EXT = new Set(['.mp4', '.mov', '.m4v', '.webm']);
const MIME = {
  '.mp4': 'video/mp4',
  '.m4v': 'video/mp4',
  '.mov': 'video/quicktime',
  '.webm': 'video/webm',
};

function safeEqual(a, b) {
  const left = Buffer.from(String(a));
  const right = Buffer.from(String(b));
  if (left.length !== right.length) return false;
  return crypto.timingSafeEqual(left, right);
}

function parseRange(header, size) {
  const match = /^bytes=(\d*)-(\d*)$/.exec(String(header || '').trim());
  if (!match) return null;
  let start = match[1] === '' ? null : Number(match[1]);
  let end = match[2] === '' ? null : Number(match[2]);
  if (start === null && end === null) return null;
  if (start === null) {
    // suffix range: last N bytes
    start = Math.max(0, size - end);
    end = size - 1;
  } else if (end === null || end >= size) {
    end = size - 1;
  }
  if (start > end || start >= size) return { invalid: true };
  return { start, end };
}

function resolveInside(root, relative) {
  const decoded = relative
    .split('/')
    .map((segment) => decodeURIComponent(segment))
    .filter((segment) => segment !== '');
  if (decoded.some((segment) => segment === '..' || segment.includes('\0') || path.isAbsolute(segment))) {
    return null;
  }
  const absolute = path.resolve(root, ...decoded);
  if (absolute !== root && !absolute.startsWith(root + path.sep)) return null;
  return absolute;
}

function encodeRelative(relativePath) {
  return relativePath
    .split(path.sep)
    .join('/')
    .split('/')
    .filter(Boolean)
    .map(encodeURIComponent)
    .join('/');
}

function createMediaServer({
  mediaDir,
  secret,
  port = 8787,
  host = '0.0.0.0',
  publicBaseUrl = '',
  allowedExtensions = VIDEO_EXT,
  logger = () => {},
} = {}) {
  if (!mediaDir) throw new Error('mediaDir لازم است.');
  if (!secret) throw new Error('media secret لازم است.');
  const root = path.resolve(mediaDir);

  const server = http.createServer((req, res) => {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    const respond = (status, body, headers = {}) => {
      res.writeHead(status, {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store',
        ...headers,
      });
      res.end(typeof body === 'string' ? body : JSON.stringify(body));
    };

    if (req.method !== 'GET' && req.method !== 'HEAD') {
      return respond(405, { error: 'متد پشتیبانی نمی‌شود' }, { Allow: 'GET, HEAD' });
    }
    if (url.pathname === '/health' || url.pathname === '/') {
      return respond(200, { ok: true, service: 'instagram-uploader-media' });
    }

    const segments = url.pathname.split('/').filter(Boolean);
    if (segments.length < 2 || !safeEqual(segments[0], secret)) {
      return respond(403, { error: 'دسترسی غیرمجاز' });
    }

    const relative = segments.slice(1).join('/');
    const file = resolveInside(root, relative);
    if (!file) return respond(400, { error: 'مسیر نامعتبر است' });
    if (!allowedExtensions.has(path.extname(file).toLowerCase())) {
      return respond(403, { error: 'فقط فایل ویدیویی قابل ارائه است' });
    }

    fs.stat(file, (err, stat) => {
      if (err || !stat.isFile()) return respond(404, { error: 'فایل پیدا نشد' });

      const headers = {
        'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
        'Accept-Ranges': 'bytes',
        'Cache-Control': 'public, max-age=300',
        'Content-Length': stat.size,
      };

      const range = parseRange(req.headers.range, stat.size);
      if (range?.invalid) {
        res.writeHead(416, { 'Content-Range': `bytes */${stat.size}` });
        return res.end();
      }
      if (range) {
        headers['Content-Length'] = range.end - range.start + 1;
        headers['Content-Range'] = `bytes ${range.start}-${range.end}/${stat.size}`;
        res.writeHead(206, headers);
        if (req.method === 'HEAD') return res.end();
        return fs.createReadStream(file, { start: range.start, end: range.end }).pipe(res);
      }

      res.writeHead(200, headers);
      if (req.method === 'HEAD') return res.end();
      fs.createReadStream(file).pipe(res);
    });
  });

  let listeningPort = null;

  function urlFor(relativePath) {
    const encoded = encodeRelative(relativePath);
    const base = (publicBaseUrl || `http://127.0.0.1:${listeningPort || port}`).replace(/\/+$/, '');
    return `${base}/${secret}/${encoded}`;
  }

  function start() {
    if (listeningPort) return Promise.resolve({ port: listeningPort, url: urlFor('') });
    return new Promise((resolve, reject) => {
      server.once('error', (error) => {
        if (error.code === 'EADDRINUSE') {
          return reject(
            new Error(
              `پورت ${port} اشغال است. اگر «serve» یا اجرای دیگری باز است آن را ببند، یا IG_MEDIA_PORT را عوض کن.`
            )
          );
        }
        reject(error);
      });
      server.listen(port, host, () => {
        server.removeAllListeners('error');
        listeningPort = server.address().port;
        logger(`سرور ویدیو روی http://${host}:${listeningPort} بالا آمد (پوشه: ${root}).`);
        resolve({ port: listeningPort, url: urlFor('') });
      });
    });
  }

  function close() {
    return new Promise((resolve) => {
      if (!listeningPort) return resolve();
      listeningPort = null;
      server.close(() => resolve());
      server.closeAllConnections?.();
    });
  }

  return { server, start, close, urlFor, get port() { return listeningPort; } };
}

module.exports = { createMediaServer, resolveInside, parseRange, encodeRelative, VIDEO_EXT };
