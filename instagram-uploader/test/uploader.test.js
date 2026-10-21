'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createMediaServer, parseRange } = require('../src/media-server');
const { loadCaptionSources, resolveCaption, trimToLimit, appendHashtags, CAPTION_LIMIT } = require('../src/captions');
const { listVideoFiles, loadState, saveState, recordPosted, recordFailure, postsInLast24h, pendingFiles, isPosted, archiveFile } = require('../src/library');
const { createGraphClient, identifyAccount, refreshLongLivedToken, GraphError } = require('../src/graph');
const { publishOnce } = require('../src/publisher');
const { loadConfig, parseHashtags } = require('../src/config');

function tempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'ig-up-'));
}

function fakeConfig(dir, extra = {}) {
  return {
    mediaDir: path.join(dir, 'media'),
    archiveDir: path.join(dir, 'media', 'posted'),
    dataDir: path.join(dir, 'data'),
    stateFile: path.join(dir, 'data', 'state.json'),
    lockFile: path.join(dir, 'data', 'lock'),
    mediaSecret: 'test-secret-123',
    mediaPort: 0,
    mediaHost: '127.0.0.1',
    publicBaseUrl: '',
    defaultCaption: '{name}',
    defaultHashtags: ['#تست'],
    dailyLimit: 10,
    maxPostsPerRun: 5,
    postIntervalSec: 0,
    archiveAfterPost: false,
    shareToFeed: true,
    dryRun: false,
    ffprobe: 'ffprobe-does-not-exist',
    graphVersion: 'v23.0',
    graphBaseUrl: 'https://graph.facebook.com',
    ...extra,
  };
}

function seedVideos(config, names = ['clip1.mp4', 'clip2.mp4']) {
  fs.mkdirSync(config.mediaDir, { recursive: true });
  for (const name of names) fs.writeFileSync(path.join(config.mediaDir, name), Buffer.alloc(2048, 7));
}

test('media server serves videos only with the secret and supports range requests', async () => {
  const dir = tempDir();
  const config = fakeConfig(dir);
  seedVideos(config, ['one.mp4']);
  fs.writeFileSync(path.join(config.mediaDir, 'secret.txt'), 'nope');
  const server = createMediaServer({ mediaDir: config.mediaDir, secret: config.mediaSecret, port: 0, host: '127.0.0.1' });
  const { port } = await server.start();
  try {
    const okUrl = server.urlFor('one.mp4');
    assert.match(okUrl, new RegExp(`:${port}/test-secret-123/one\\.mp4$`));

    const full = await fetch(okUrl);
    assert.equal(full.status, 200);
    assert.equal(full.headers.get('content-type'), 'video/mp4');
    assert.equal(full.headers.get('accept-ranges'), 'bytes');
    assert.equal((await full.arrayBuffer()).byteLength, 2048);

    const ranged = await fetch(okUrl, { headers: { Range: 'bytes=100-199' } });
    assert.equal(ranged.status, 206);
    assert.equal(ranged.headers.get('content-range'), 'bytes 100-199/2048');
    assert.equal((await ranged.arrayBuffer()).byteLength, 100);

    const head = await fetch(okUrl, { method: 'HEAD' });
    assert.equal(head.status, 200);
    assert.equal(head.headers.get('content-length'), '2048');

    const wrongSecret = await fetch(`http://127.0.0.1:${port}/wrong/one.mp4`);
    assert.equal(wrongSecret.status, 403);

    const traversal = await fetch(`http://127.0.0.1:${port}/test-secret-123/..%2f..%2fetc%2fpasswd`);
    assert.ok([400, 403, 404].includes(traversal.status));

    const nonVideo = await fetch(`http://127.0.0.1:${port}/test-secret-123/secret.txt`);
    assert.equal(nonVideo.status, 403);

    const missing = await fetch(`http://127.0.0.1:${port}/test-secret-123/none.mp4`);
    assert.equal(missing.status, 404);
  } finally {
    await server.close();
  }
});

test('parseRange handles suffix, open-ended and invalid headers', () => {
  assert.deepEqual(parseRange('bytes=0-99', 1000), { start: 0, end: 99 });
  assert.deepEqual(parseRange('bytes=500-', 1000), { start: 500, end: 999 });
  assert.deepEqual(parseRange('bytes=-100', 1000), { start: 900, end: 999 });
  assert.deepEqual(parseRange('bytes=900-5000', 1000), { start: 900, end: 999 });
  assert.deepEqual(parseRange('bytes=1000-1200', 1000), { invalid: true });
  assert.equal(parseRange('items=0-5', 1000), null);
});

test('captions.json, sidecar .txt and default template resolve in priority order', () => {
  const dir = tempDir();
  const config = fakeConfig(dir, { defaultHashtags: [] });
  seedVideos(config, ['a.mp4', 'b.mp4', 'c.mp4', 'd.mp4']);
  fs.writeFileSync(path.join(config.mediaDir, 'b.txt'), 'کپشن از فایل متنی');
  fs.writeFileSync(
    path.join(config.mediaDir, 'captions.json'),
    JSON.stringify({ 'a.mp4': 'کپشن از جیسون', 'c.mp4': { caption: 'با کاور', cover: 'https://x/c.jpg' } })
  );
  const sources = loadCaptionSources(config.mediaDir);
  assert.equal(resolveCaption({ relativePath: 'a.mp4', mediaDir: config.mediaDir, config, sources }).caption, 'کپشن از جیسون');
  assert.equal(resolveCaption({ relativePath: 'b.mp4', mediaDir: config.mediaDir, config, sources }).caption, 'کپشن از فایل متنی');
  const c = resolveCaption({ relativePath: 'c.mp4', mediaDir: config.mediaDir, config, sources });
  assert.match(c.caption, /با کاور/);
  assert.equal(c.cover, 'https://x/c.jpg');
  assert.match(resolveCaption({ relativePath: 'd.mp4', mediaDir: config.mediaDir, config, sources }).caption, /^d/);
});

test('caption template placeholders and limits work', () => {
  const dir = tempDir();
  const config = fakeConfig(dir, { defaultCaption: '🎬 {name} — {date}' });
  fs.mkdirSync(config.mediaDir, { recursive: true });
  fs.writeFileSync(path.join(config.mediaDir, '01 - my clip.mp4'), 'x');
  const sources = loadCaptionSources(config.mediaDir);
  const { caption } = resolveCaption({ relativePath: '01 - my clip.mp4', mediaDir: config.mediaDir, config, sources, date: new Date('2026-10-21T00:00:00Z') });
  assert.match(caption, /^🎬 my clip — 2026-10-21/);
  assert.match(caption, /#تست/);

  assert.ok(trimToLimit('a '.repeat(2000)).length <= CAPTION_LIMIT);
  const hashtags = Array.from({ length: 40 }, (_, i) => `#t${i}`);
  assert.equal((appendHashtags('سلام', hashtags).match(/#/g) || []).length, 30);
  assert.deepEqual(parseHashtags('a #b, #c'), ['#a', '#b', '#c']);
});

test('library scans folders, ignores archive dir, and tracks posted files', () => {
  const dir = tempDir();
  const config = fakeConfig(dir);
  seedVideos(config, ['b.mp4', 'a10.mp4', 'a2.mp4']);
  fs.mkdirSync(path.join(config.mediaDir, 'sub'), { recursive: true });
  fs.writeFileSync(path.join(config.mediaDir, 'sub', 'nested.mov'), 'x');
  fs.writeFileSync(path.join(config.mediaDir, 'notes.txt'), 'x');
  fs.mkdirSync(config.archiveDir, { recursive: true });
  fs.writeFileSync(path.join(config.archiveDir, 'old.mp4'), 'x');

  const files = listVideoFiles(config.mediaDir, { excludeDirs: [config.archiveDir] });
  const names = files.map((f) => f.relative);
  assert.deepEqual(names, ['a2.mp4', 'a10.mp4', 'b.mp4', path.join('sub', 'nested.mov')]);

  const state = loadState(config.stateFile);
  assert.equal(state.posted && Object.keys(state.posted).length, 0);
  assert.equal(pendingFiles(files, state).length, 4);
  recordPosted(state, files[0], { mediaId: 'm1', containerId: 'c1', caption: 'x' });
  saveState(config.stateFile, state);
  const reloaded = loadState(config.stateFile);
  assert.equal(isPosted(reloaded, files[0]), true);
  assert.equal(pendingFiles(files, reloaded).length, 3);
  assert.equal(postsInLast24h(reloaded), 1);
  recordFailure(reloaded, files[1], 'boom');
  assert.equal(reloaded.attempts[files[1].relative].count, 1);
});

test('archive moves files into the posted folder without overwriting', () => {
  const dir = tempDir();
  const config = fakeConfig(dir);
  seedVideos(config, ['clip.mp4']);
  const files = listVideoFiles(config.mediaDir);
  const target = archiveFile(files[0], config.archiveDir);
  assert.ok(fs.existsSync(target));
  assert.equal(fs.existsSync(files[0].absolute), false);
  fs.writeFileSync(path.join(config.mediaDir, 'clip.mp4'), 'again');
  const second = archiveFile(listVideoFiles(config.mediaDir)[0], config.archiveDir);
  assert.notEqual(second, target);
  assert.equal(fs.readdirSync(config.archiveDir).length, 2);
});

test('graph client builds the official reels flow and surfaces API errors', async () => {
  const calls = [];
  const responses = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url: String(url), init });
    const next = responses.shift() || { status: 200, body: {} };
    return new Response(JSON.stringify(next.body), {
      status: next.status,
      headers: { 'content-type': 'application/json' },
    });
  };
  const client = createGraphClient({
    accessToken: 'TOKEN',
    igUserId: '17841400000000000',
    baseUrl: 'https://graph.facebook.com',
    version: 'v23.0',
    fetchImpl,
  });

  responses.push({ status: 200, body: { id: 'container-1' } });
  const containerId = await client.createReelContainer({ videoUrl: 'https://cdn.example/v.mp4', caption: 'hi' });
  assert.equal(containerId, 'container-1');
  const createCall = calls[0];
  assert.equal(createCall.url, 'https://graph.facebook.com/v23.0/17841400000000000/media');
  assert.equal(createCall.init.method, 'POST');
  const body = new URLSearchParams(createCall.init.body.toString());
  assert.equal(body.get('media_type'), 'REELS');
  assert.equal(body.get('video_url'), 'https://cdn.example/v.mp4');
  assert.equal(body.get('caption'), 'hi');
  assert.equal(body.get('access_token'), 'TOKEN');

  responses.push({ status: 200, body: { status_code: 'FINISHED' } });
  const status = await client.waitForContainer('container-1', { intervalMs: 1 });
  assert.equal(status.status_code, 'FINISHED');
  assert.match(calls[1].url, /\/container-1\?fields=status_code/);

  responses.push({ status: 200, body: { id: 'media-9' } });
  assert.equal(await client.publish('container-1'), 'media-9');
  assert.equal(calls[2].url, 'https://graph.facebook.com/v23.0/17841400000000000/media_publish');

  responses.push({ status: 400, body: { error: { message: 'Invalid OAuth access token', code: 190 } } });
  await assert.rejects(() => client.getAccount(), (error) => {
    assert.ok(error instanceof GraphError);
    assert.equal(error.code, 190);
    assert.match(error.message, /خطای اینستاگرام/);
    return true;
  });
});

test('waitForContainer fails fast when Instagram reports ERROR', async () => {
  const fetchImpl = async () => new Response(JSON.stringify({ status_code: 'ERROR', status: 'bad file' }), { status: 200 });
  const client = createGraphClient({ accessToken: 't', igUserId: '1', fetchImpl });
  await assert.rejects(() => client.waitForContainer('c', { intervalMs: 1 }), /شکست خورد/);
});

test('identifyAccount resolves the account id for both API bases', async () => {
  const seen = [];
  const igFetch = async (url) => {
    seen.push(String(url));
    return new Response(JSON.stringify({ id: '17841400000000000', user_id: '17841400000000000', username: 'my.page', account_type: 'BUSINESS' }), { status: 200 });
  };
  const ig = await identifyAccount({ accessToken: 'IGQ_x', baseUrl: 'https://graph.instagram.com', fetchImpl: igFetch });
  assert.deepEqual(ig, { id: '17841400000000000', username: 'my.page', accountType: 'BUSINESS', via: 'instagram-login' });
  assert.match(seen[0], /graph\.instagram\.com\/v23\.0\/me\?fields=.*username/);

  const fbFetch = async (url) => {
    const href = String(url);
    if (href.includes('/me/accounts')) {
      return new Response(JSON.stringify({ data: [{ name: 'Page', instagram_business_account: { id: '999', username: 'via_page' } }] }), { status: 200 });
    }
    return new Response(JSON.stringify({ error: { message: 'unknown path' } }), { status: 400 });
  };
  const fb = await identifyAccount({ accessToken: 'EAA_x', baseUrl: 'https://graph.facebook.com', fetchImpl: fbFetch });
  assert.equal(fb.id, '999');
  assert.equal(fb.via, 'facebook-login');
});

test('publishing limit and token refresh read the documented endpoints', async () => {
  const urls = [];
  const fetchImpl = async (url) => {
    urls.push(String(url));
    if (String(url).includes('content_publishing_limit')) {
      return new Response(JSON.stringify({ data: [{ quota_usage: 7, config: { quota_total: 50, quota_duration: 86400 } }] }), { status: 200 });
    }
    return new Response(JSON.stringify({ access_token: 'NEW_TOKEN', expires_in: 5184000, token_type: 'bearer' }), { status: 200 });
  };
  const client = createGraphClient({ accessToken: 't', igUserId: '42', fetchImpl });
  const limit = await client.getPublishingLimit();
  assert.deepEqual(limit, { quotaTotal: 50, quotaUsage: 7, durationSec: 86400 });
  assert.match(urls[0], /\/42\/content_publishing_limit\?fields=config%2Cquota_usage/);

  const refreshed = await refreshLongLivedToken({ accessToken: 'OLD', fetchImpl, baseUrl: 'https://graph.instagram.com/' });
  assert.equal(refreshed.accessToken, 'NEW_TOKEN');
  assert.equal(refreshed.expiresInSec, 5184000);
  assert.match(urls[1], /^https:\/\/graph\.instagram\.com\/refresh_access_token\?grant_type=ig_refresh_token&access_token=OLD$/);
});

test('publishOnce publishes new files once, records state, and skips processed ones', async () => {
  const dir = tempDir();
  const config = fakeConfig(dir);
  seedVideos(config, ['v1.mp4', 'v2.mp4']);
  fs.writeFileSync(path.join(config.mediaDir, 'v1.txt'), 'کپشن اول');
  let counter = 0;
  const created = [];
  const client = {
    async createReelContainer(payload) {
      created.push(payload);
      counter += 1;
      return `container-${counter}`;
    },
    async waitForContainer(id) {
      return { status_code: 'FINISHED', id };
    },
    async publish(id) {
      return `media-${id}`;
    },
  };
  const state = loadState(config.stateFile);
  const logs = [];
  const summary = await publishOnce({ config, client, state, logger: (line) => logs.push(line), probe: false });

  assert.equal(summary.published.length, 2);
  assert.deepEqual(created.map((c) => c.caption), ['کپشن اول\n\n#تست', 'v2\n\n#تست']);
  assert.match(created[0].videoUrl, /^http:\/\/127\.0\.0\.1:\d+\/test-secret-123\/v1\.mp4$/);
  assert.equal(state.posted['v1.mp4'].mediaId, 'media-container-1');
  assert.equal(Object.keys(state.posted).length, 2);

  const secondRun = await publishOnce({ config, client, state, logger: () => {}, probe: false });
  assert.equal(secondRun.published.length, 0);
  assert.equal(secondRun.skipped.length, 2);
  saveState(config.stateFile, state);
  assert.equal(loadState(config.stateFile).posted['v2.mp4'].mediaId, 'media-container-2');
});

test('dry run never touches the network or the state ledger', async () => {
  const dir = tempDir();
  const config = fakeConfig(dir);
  seedVideos(config, ['v1.mp4']);
  const state = loadState(config.stateFile);
  const summary = await publishOnce({ config, client: {}, state, dryRun: true, logger: () => {}, probe: false });
  assert.equal(summary.published.length, 1);
  assert.equal(summary.published[0].dryRun, true);
  assert.equal(Object.keys(state.posted).length, 0);
});

test('a failing upload is recorded, does not block the next file, and archive keeps the file', async () => {
  const dir = tempDir();
  const config = fakeConfig(dir, { archiveAfterPost: true });
  seedVideos(config, ['bad.mp4', 'good.mp4']);
  const client = {
    async createReelContainer({ videoUrl }) {
      if (videoUrl.endsWith('bad.mp4')) throw new Error('invalid video format');
      return 'ok-container';
    },
    async waitForContainer(id) {
      return { status_code: 'FINISHED', id };
    },
    async publish() {
      return 'ok-media';
    },
  };
  const state = loadState(config.stateFile);
  const summary = await publishOnce({ config, client, state, logger: () => {}, probe: false });
  assert.deepEqual(summary.failed.map((f) => f.file), ['bad.mp4']);
  assert.deepEqual(summary.published.map((p) => p.file), ['good.mp4']);
  assert.equal(state.attempts['bad.mp4'].count, 1);
  assert.ok(fs.existsSync(path.join(config.archiveDir, 'good.mp4')));
  assert.ok(fs.existsSync(path.join(config.mediaDir, 'bad.mp4')));
});

test('daily limit stops publishing', async () => {
  const dir = tempDir();
  const config = fakeConfig(dir, { dailyLimit: 1 });
  seedVideos(config, ['v1.mp4', 'v2.mp4']);
  const state = loadState(config.stateFile);
  recordPosted(state, { relative: 'already.mp4', size: 1, mtimeMs: 1 }, { mediaId: 'm', containerId: 'c', caption: 'x' });
  const summary = await publishOnce({ config, client: {}, state, logger: () => {}, probe: false });
  assert.equal(summary.published.length, 0);
});

test('loadConfig reads env, creates a media secret, and validates required settings', () => {
  const dir = tempDir();
  const previous = { ...process.env };
  Object.assign(process.env, {
    IG_DATA_DIR: path.join(dir, 'data'),
    IG_MEDIA_DIR: path.join(dir, 'media'),
    IG_DAILY_LIMIT: '7',
    IG_DEFAULT_HASHTAGS: '#one two',
  });
  delete process.env.IG_MEDIA_SECRET;
  try {
    const config = loadConfig();
    assert.equal(config.mediaDir, path.join(dir, 'media'));
    assert.equal(config.dailyLimit, 7);
    assert.deepEqual(config.defaultHashtags, ['#one', '#two']);
    assert.ok(config.mediaSecret.length >= 16);
    assert.ok(fs.existsSync(path.join(dir, 'data', 'media-secret.txt')));
    const { validateForPublishing } = require('../src/config');
    const problems = validateForPublishing(config);
    assert.equal(problems.length, 3);
  } finally {
    for (const key of Object.keys(process.env)) if (!(key in previous)) delete process.env[key];
    Object.assign(process.env, previous);
  }
});
