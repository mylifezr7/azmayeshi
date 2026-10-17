const DAY = 86400;

class DemoProvider {
  async getCandles({ symbol = 'AAPL', timeframe = '15m', limit = 260 }) {
    const ticker = String(symbol).toUpperCase();
    let seed = [...ticker].reduce((n, c) => (n * 31 + c.charCodeAt(0)) >>> 0, 2026);
    const rand = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
    const bases = { AAPL: 225, MSFT: 510, NVDA: 145, TSLA: 330, AMZN: 220, GOOGL: 180, META: 700, SPY: 660, QQQ: 590, AMD: 160 };
    let close = bases[ticker] || 100 + rand() * 400;
    const candles = [];
    const intervalMs = { '5m': 300000, '15m': 900000, '30m': 1800000, '1h': 3600000, '1d': 86400000 }[timeframe] || 900000;
    for (let i = 0; i < limit; i++) {
      const drift = Math.sin(i / 35) * .0008;
      const open = close; close = Math.max(.1, close * (1 + drift + (rand() - .505) * .025));
      const span = close * (.003 + rand() * .018);
      candles.push({ time: Date.now() - (limit - i) * intervalMs, open, high: Math.max(open, close) + span * rand(), low: Math.max(.01, Math.min(open, close) - span * rand()), close, volume: 100000 + rand() * 3000000 });
    }
    return candles;
  }
  async getFundamentals() { return null; }
  async getNews() { return []; }
  async getQuote({ candles }) { const price = candles.at(-1).close, previousClose = candles.at(-2).close; return { price, previousClose, changePct: (price / previousClose - 1) * 100, changeLabel: 'از کندل قبلی · مصنوعی', source: 'دادهٔ مصنوعی' }; }
}

class FinnhubProvider {
  constructor(token = process.env.FINNHUB_API_KEY) {
    if (!token) throw new Error('برای دریافت دادهٔ زندهٔ سهام آمریکا، ابتدا FINNHUB_API_KEY را در متغیرهای محیطی تنظیم کنید.');
    this.token = token;
    this.base = 'https://finnhub.io/api/v1';
  }
  async request(path) {
    const response = await fetch(`${this.base}${path}${path.includes('?') ? '&' : '?'}token=${encodeURIComponent(this.token)}`, { signal: AbortSignal.timeout(12000) });
    if (!response.ok) throw new Error(`منبع Finnhub پاسخ ${response.status} داد.`);
    return response.json();
  }
  async getCandles({ symbol = 'AAPL', timeframe = '1D', limit = 260 }) {
    const resolutions = { '5m': '5', '15m': '15', '30m': '30', '1h': '60', '1d': 'D' };
    const resolution = resolutions[timeframe] || 'D';
    const now = Math.floor(Date.now() / 1000);
    const barsPerDay = { '5': 78, '15': 26, '30': 13, '60': 7 }[resolution];
    const span = resolution === 'D' ? Math.min(730, limit * 2) : Math.min(30, Math.ceil(limit / barsPerDay) + 3);
    const query = new URLSearchParams({ symbol: symbol.toUpperCase(), resolution, from: String(now - span * DAY), to: String(now) });
    const result = await this.request(`/stock/candle?${query}`);
    if (result.s !== 'ok' || !Array.isArray(result.c)) throw new Error('Finnhub برای این نماد کندلی برنگرداند. نماد را با پسوند بازار بررسی کنید.');
    const start = Math.max(0, result.c.length - limit);
    return result.c.slice(start).map((close, i) => { const j = start + i; return { time: result.t[j] * 1000, open: result.o[j], high: result.h[j], low: result.l[j], close, volume: result.v[j] }; });
  }
  async getQuote({ symbol }) {
    const q = await this.request(`/quote?symbol=${encodeURIComponent(symbol.toUpperCase())}`);
    if (!q.c) throw new Error('قیمت جاری برای نماد پیدا نشد.');
    return { price: q.c, previousClose: q.pc, changePct: q.dp, changeLabel: 'روزانه', source: 'Finnhub', updatedAt: q.t ? new Date(q.t * 1000).toISOString() : null };
  }
  async getFundamentals({ symbol }) {
    const [profile, response] = await Promise.all([
      this.request(`/stock/profile2?symbol=${encodeURIComponent(symbol.toUpperCase())}`),
      this.request(`/stock/metric?symbol=${encodeURIComponent(symbol.toUpperCase())}&metric=all`)
    ]);
    const m = response.metric || {};
    const pick = (...keys) => { for (const key of keys) if (m[key] !== undefined && m[key] !== null && Number.isFinite(Number(m[key]))) return Number(m[key]); return null; };
    return { name: profile.name || symbol.toUpperCase(), exchange: profile.exchange || 'US', marketCap: pick('marketCapitalization'), pe: pick('peTTM', 'peBasicExclExtraTTM'), eps: pick('epsTTM'), revenueGrowth: pick('revenueGrowthTTMYoy'), debtToEquity: pick('totalDebt/totalEquityQuarterly', 'totalDebt/totalEquityAnnual'), beta: pick('beta'), roe: pick('roeTTM'), dividendYield: pick('dividendYieldIndicatedAnnual'), currency: profile.currency || 'USD', source: 'Finnhub' };
  }
  async getNews({ symbol }) {
    const to = new Date(), from = new Date(Date.now() - 7 * DAY * 1000);
    const dates = d => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
    const items = await this.request(`/company-news?symbol=${encodeURIComponent(symbol.toUpperCase())}&from=${dates(from)}&to=${dates(to)}`);
    return (Array.isArray(items) ? items : []).slice(0, 8).map(x => ({ headline: x.headline, source: x.source, url: x.url, time: x.datetime ? new Date(x.datetime * 1000).toISOString() : null, sentiment: scoreHeadline(x.headline || '') }));
  }
}
function scoreHeadline(text) {
  const positive = ['beat', 'upgrade', 'growth', 'surge', 'profit', 'record', 'approval', 'strong', 'rise', 'gain', 'bullish'];
  const negative = ['miss', 'downgrade', 'loss', 'fall', 'drop', 'lawsuit', 'probe', 'cut', 'weak', 'bearish', 'warning'];
  const value = text.toLowerCase();
  const score = positive.filter(w => value.includes(w)).length - negative.filter(w => value.includes(w)).length;
  return score > 0 ? 'مثبتِ واژه‌محور' : score < 0 ? 'منفیِ واژه‌محور' : 'خنثی';
}
function createProvider(id = 'demo') {
  if (id === 'demo') return new DemoProvider();
  if (id === 'finnhub') return new FinnhubProvider();
  throw new Error('منبع داده ناشناخته است.');
}
module.exports = { createProvider, DemoProvider, FinnhubProvider, scoreHeadline };
