const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { URL } = require('node:url');
const { createProvider } = require('./src/providers');
const { analyze, marketClock } = require('./src/analysis');

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 3000);
function json(res, status, body) { res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(body)); }
function serveStatic(res, pathname) {
  const requested = pathname === '/' ? '/public/index.html' : `/public${pathname}`;
  const file = path.resolve(ROOT, `.${requested}`);
  if (!file.startsWith(path.join(ROOT, 'public') + path.sep)) return json(res, 403, { error: 'دسترسی غیرمجاز' });
  fs.readFile(file, (err, content) => {
    if (err) return json(res, 404, { error: 'فایل پیدا نشد' });
    const ext = path.extname(file);
    const type = ext === '.css' ? 'text/css' : ext === '.js' ? 'text/javascript' : 'text/html';
    res.writeHead(200, { 'Content-Type': `${type}; charset=utf-8` }); res.end(content);
  });
}
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  try {
    if (req.method === 'GET' && url.pathname === '/api/health') return json(res, 200, { ok: true, product: 'analysis-only', realOrders: false, fundamentals: Boolean(process.env.FINNHUB_API_KEY) });
    if (req.method === 'GET' && url.pathname === '/api/providers') return json(res, 200, [
      { id: 'demo', name: 'دمو', available: true, description: 'کندل مصنوعی؛ برای بررسی عملکرد رابط و محاسبات' },
      { id: 'finnhub', name: 'Finnhub', available: Boolean(process.env.FINNHUB_API_KEY), description: 'کندل، قیمت، مشخصات بنیادی و خبر با کلید API شخصی' }
    ]);
    if (req.method === 'GET' && url.pathname === '/api/market-clock') return json(res, 200, marketClock());
    if (req.method === 'GET' && url.pathname === '/api/analysis') {
      const symbol = (url.searchParams.get('symbol') || 'AAPL').trim().toUpperCase();
      if (!/^[A-Z][A-Z0-9.\-]{0,11}$/.test(symbol)) return json(res, 400, { error: 'نماد معتبر نیست.' });
      const timeframe = url.searchParams.get('timeframe') || '15m';
      if (!['5m','15m','30m','1h','1d'].includes(timeframe)) return json(res, 400, { error: 'بازهٔ زمانی پشتیبانی نمی‌شود.' });
      const providerId = url.searchParams.get('provider') || 'demo';
      const provider = createProvider(providerId);
      const candles = await provider.getCandles({ symbol, timeframe, limit: 260 });
      const [quoteResult, fundamentalsResult, newsResult] = await Promise.allSettled([
        provider.getQuote({ symbol, candles }), provider.getFundamentals({ symbol }), provider.getNews({ symbol })
      ]);
      const quote = quoteResult.status === 'fulfilled' ? quoteResult.value : { price: candles.at(-1).close, previousClose: candles.at(-2)?.close, source: `${providerId} · قیمت کندل` };
      const fundamentals = fundamentalsResult.status === 'fulfilled' ? fundamentalsResult.value : null;
      const news = newsResult.status === 'fulfilled' ? newsResult.value : [];
      const result = analyze({ symbol, candles, quote, fundamentals, news, riskPct: url.searchParams.get('risk'), balance: url.searchParams.get('balance'), timeframe });
      result.dataStatus = {
        provider: providerId,
        quote: quoteResult.status === 'fulfilled' ? 'وصل' : 'استفاده از آخرین کندل',
        fundamentals: fundamentals ? 'وصل' : 'دردسترس نیست',
        news: news.length ? 'وصل' : 'خبر تازه/دردسترس نیست',
        fundamentalsError: fundamentalsResult.status === 'rejected' ? fundamentalsResult.reason.message : null,
        newsError: newsResult.status === 'rejected' ? newsResult.reason.message : null
      };
      return json(res, 200, result);
    }
    if (req.method === 'GET') return serveStatic(res, url.pathname);
    return json(res, 404, { error: 'مسیر پیدا نشد' });
  } catch (error) { return json(res, error.status || 400, { error: error.message || 'خطای غیرمنتظره' }); }
});
server.listen(PORT, '0.0.0.0', () => console.log(`Market analysis listening on 0.0.0.0:${PORT}`));
