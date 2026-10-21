'use strict';
// Minimal client for the official Instagram Graph API content-publishing flow:
//   1) POST /{ig-user-id}/media            -> creation_id  (container)
//   2) GET  /{creation-id}?fields=status_code  until FINISHED
//   3) POST /{ig-user-id}/media_publish    -> published media id
// No cookie/session automation, no password login: only the documented API.

const DEFAULT_TIMEOUT_MS = 60_000;

class GraphError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = 'GraphError';
    Object.assign(this, details);
  }
}

function redact(text) {
  return String(text).replace(/(access_token|access-token)=([^&\s]+)/gi, '$1=***');
}

function sleep(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) return reject(new Error('aborted'));
    const timer = setTimeout(() => {
      signal?.removeEventListener?.('abort', onAbort);
      resolve();
    }, ms);
    function onAbort() {
      clearTimeout(timer);
      reject(new Error('aborted'));
    }
    signal?.addEventListener?.('abort', onAbort, { once: true });
  });
}

function createGraphClient({
  accessToken,
  igUserId,
  version = 'v23.0',
  baseUrl = 'https://graph.facebook.com',
  fetchImpl = globalThis.fetch,
  timeoutMs = DEFAULT_TIMEOUT_MS,
  logger = () => {},
} = {}) {
  if (typeof fetchImpl !== 'function') throw new Error('fetch در دسترس نیست؛ Node.js 18 یا جدیدتر لازم است.');
  if (!accessToken) throw new Error('accessToken لازم است.');
  if (!igUserId) throw new Error('igUserId لازم است.');

  let lastUsage = null;

  async function request(method, endpoint, params = {}) {
    const url = new URL(`${baseUrl}/${version}/${String(endpoint).replace(/^\/+/, '')}`);
    const body = new URLSearchParams();
    for (const [key, value] of Object.entries(params)) {
      if (value === undefined || value === null) continue;
      body.set(key, String(value));
    }
    body.set('access_token', accessToken);

    const init = { method, signal: AbortSignal.timeout(timeoutMs) };
    if (method === 'GET' || method === 'DELETE') {
      for (const [key, value] of body) url.searchParams.set(key, value);
    } else {
      init.body = body;
      init.headers = { 'Content-Type': 'application/x-www-form-urlencoded' };
    }

    let response;
    try {
      response = await fetchImpl(url, init);
    } catch (error) {
      throw new GraphError(`ارتباط با Graph API برقرار نشد: ${error.message}`, { cause: error });
    }

    const usageHeader = response.headers?.get?.('x-app-usage');
    if (usageHeader) {
      try {
        lastUsage = JSON.parse(usageHeader);
      } catch {
        /* header is best-effort */
      }
    }

    let payload = null;
    const text = await response.text();
    if (text) {
      try {
        payload = JSON.parse(text);
      } catch {
        payload = { raw: text };
      }
    }

    if (!response.ok || payload?.error) {
      const apiError = payload?.error || {};
      const message = apiError.error_user_msg || apiError.message || payload?.raw || `HTTP ${response.status}`;
      throw new GraphError(`خطای اینستاگرام: ${redact(message)}`, {
        status: response.status,
        code: apiError.code,
        subcode: apiError.error_subcode,
        type: apiError.type,
        fbtraceId: apiError.fbtrace_id,
      });
    }
    return payload || {};
  }

  return {
    get lastUsage() {
      return lastUsage;
    },

    async getAccount() {
      return request('GET', igUserId, {
        fields: 'id,username,name,account_type,media_count,followers_count',
      });
    },

    // Authoritative source for how many API posts the account may still publish
    // in the rolling 24h window (Meta keeps changing the documented number).
    async getPublishingLimit() {
      const payload = await request('GET', `${igUserId}/content_publishing_limit`, {
        fields: 'config,quota_usage',
      });
      const entry = (payload.data || [])[0] || {};
      const total = Number(entry.config?.quota_total);
      const usage = Number(entry.quota_usage);
      return {
        quotaTotal: Number.isFinite(total) ? total : null,
        quotaUsage: Number.isFinite(usage) ? usage : null,
        durationSec: Number(entry.config?.quota_duration) || null,
      };
    },

    async getRecentMedia(limit = 10) {
      const data = await request('GET', `${igUserId}/media`, {
        fields: 'id,caption,timestamp,media_type,permalink',
        limit,
      });
      return data.data || [];
    },

    // Reels are the supported video surface for API publishing.
    async createReelContainer({ videoUrl, caption, shareToFeed = true, coverUrl, thumbOffset } = {}) {
      if (!videoUrl) throw new Error('videoUrl لازم است (URL عمومی HTTPS فایل ویدیو).');
      const payload = await request('POST', `${igUserId}/media`, {
        media_type: 'REELS',
        video_url: videoUrl,
        caption,
        share_to_feed: shareToFeed ? 'true' : 'false',
        cover_url: coverUrl,
        thumb_offset: thumbOffset,
      });
      if (!payload.id) throw new GraphError('اینستاگرام شناسهٔ container برنگرداند.', { body: payload });
      return payload.id;
    },

    async getContainerStatus(containerId) {
      return request('GET', containerId, { fields: 'status_code,status,id' });
    },

    // Polls until the video is processed. Reels can take several minutes for big files.
    async waitForContainer(containerId, { timeoutMs: waitMs = 15 * 60_000, intervalMs = 10_000, onStatus } = {}) {
      const startedAt = Date.now();
      let attempt = 0;
      while (true) {
        attempt += 1;
        const status = await this.getContainerStatus(containerId);
        const code = status.status_code || 'UNKNOWN';
        onStatus?.({ attempt, code, status: status.status });
        logger?.(`وضعیت پردازش: ${code}${status.status ? ` — ${redact(status.status)}` : ''}`);
        if (code === 'FINISHED' || code === 'PUBLISHED') return status;
        if (code === 'ERROR' || code === 'EXPIRED') {
          throw new GraphError(
            `پردازش ویدیو در اینستاگرام شکست خورد (${code}): ${status.status || 'بدون توضیح'}`,
            { containerId, statusCode: code }
          );
        }
        if (Date.now() - startedAt > waitMs) {
          throw new GraphError(`پردازش ویدیو بیش از ${Math.round(waitMs / 60000)} دقیقه طول کشید.`, {
            containerId,
            timeout: true,
          });
        }
        await sleep(intervalMs);
      }
    },

    async publish(containerId) {
      const payload = await request('POST', `${igUserId}/media_publish`, {
        creation_id: containerId,
      });
      if (!payload.id) throw new GraphError('اینستاگرام شناسهٔ پست منتشرشده را برنگرداند.', { body: payload });
      return payload.id;
    },
  };
}

// Instagram-Login long-lived tokens can be extended by 60 days when they are at
// least 24 hours old and not yet expired. Requires no app secret.
async function refreshLongLivedToken({
  accessToken,
  baseUrl = 'https://graph.instagram.com',
  fetchImpl = globalThis.fetch,
} = {}) {
  if (!accessToken) throw new Error('accessToken لازم است.');
  const url = new URL(`${baseUrl.replace(/\/+$/, '')}/refresh_access_token`);
  url.searchParams.set('grant_type', 'ig_refresh_token');
  url.searchParams.set('access_token', accessToken);
  const response = await fetchImpl(url, { signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS) });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload.error) {
    const info = payload.error || {};
    throw new GraphError(`تمدید توکن ناموفق بود: ${redact(info.message || `HTTP ${response.status}`)}`, {
      code: info.code,
    });
  }
  return {
    accessToken: payload.access_token,
    expiresInSec: Number(payload.expires_in) || null,
    tokenType: payload.token_type || 'bearer',
  };
}

module.exports = { createGraphClient, refreshLongLivedToken, GraphError, sleep, redact };
