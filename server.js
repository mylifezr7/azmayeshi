const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { URL } = require('node:url');
const { strategies, addStrategy, getStrategy } = require('./src/strategies');
const { runBacktest } = require('./src/backtest');
const { createProvider } = require('./src/providers');
const { analyzeMarket } = require('./src/intelligence');
const { optimize } = require('./src/optimizer');
const learningStore = require('./src/learning');

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 3000);
let paperSession = null;

function json(res, status, body) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(body));
}
function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => { data += chunk; if (data.length > 1_000_000) reject(new Error('درخواست بیش از حد بزرگ است')); });
    req.on('end', () => { try { resolve(data ? JSON.parse(data) : {}); } catch { reject(new Error('JSON نامعتبر است')); } });
    req.on('error', reject);
  });
}
function serveStatic(req, res, pathname) {
  const requested = pathname === '/' ? '/public/index.html' : `/public${pathname}`;
  const file = path.resolve(ROOT, `.${requested}`);
  if (!file.startsWith(path.join(ROOT, 'public'))) return json(res, 403, { error: 'دسترسی غیرمجاز' });
  fs.readFile(file, (err, content) => {
    if (err) return json(res, 404, { error: 'فایل پیدا نشد' });
    const type = file.endsWith('.css') ? 'text/css' : file.endsWith('.js') ? 'text/javascript' : 'text/html';
    res.writeHead(200, { 'Content-Type': `${type}; charset=utf-8` }); res.end(content);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  try {
    if (req.method === 'GET' && url.pathname === '/api/health') return json(res, 200, { ok: true, mode: 'paper-only', providers: ['mock', 'binance', 'soft4fx'], capabilities: ['analysis', 'strategy-scan', 'risk-controls', 'parameter-search'] });
    if (req.method === 'GET' && url.pathname === '/api/learning') return json(res, 200, learningStore.summary());
    if (req.method === 'GET' && url.pathname === '/api/intelligence') {
      const symbol = url.searchParams.get('symbol') || 'BTCUSDT';
      const provider = createProvider(url.searchParams.get('provider') || 'mock');
      const candles = await provider.getCandles({ symbol, timeframe: url.searchParams.get('timeframe') || '1h', limit: 700 });
      return json(res, 200, analyzeMarket({ candles, strategies: strategies.list(), requestedRisk: url.searchParams.get('risk'), learning: learningStore.summary() }));
    }
    if (req.method === 'POST' && url.pathname === '/api/scan') {
      const body = await readBody(req); const providerId = body.provider || 'mock';
      const symbols = (Array.isArray(body.symbols) ? body.symbols : ['BTCUSDT','ETHUSDT','SOLUSDT','BNBUSDT','XRPUSDT','ADAUSDT']).slice(0, 12);
      const rows = await Promise.all(symbols.map(async symbol => {
        const candles = await createProvider(providerId).getCandles({ symbol, timeframe: body.timeframe || '1h', limit: 300 });
        const analysis = analyzeMarket({ candles, strategies: strategies.list(), requestedRisk: body.risk, learning: learningStore.summary() });
        return { symbol, price: analysis.symbolPrice, direction: analysis.decision.direction, confidencePct: analysis.decision.confidencePct, agreementPct: analysis.decision.agreementPct, trend: analysis.trend, momentum20Pct: analysis.momentum20Pct, liquidityScore: analysis.liquidity.score };
      }));
      return json(res, 200, { results: rows.sort((a, b) => b.confidencePct - a.confidencePct), note: providerId === 'mock' ? 'نمادها با داده مصنوعی اسکن شده‌اند.' : 'رتبه‌بندی اکتشافی است؛ هیچ توصیه یا اجرای معامله‌ای انجام نمی‌شود.' });
    }
    if (req.method === 'POST' && url.pathname === '/api/optimize') {
      const body = await readBody(req); const strategy = getStrategy(body.strategyId);
      const candles = await createProvider(body.provider || 'mock').getCandles({ symbol: body.symbol || 'BTCUSDT', timeframe: body.timeframe || '1h', limit: 700 });
      return json(res, 200, optimize({ strategy, candles, risk: body.risk, initialBalance: body.balance }));
    }
    if (req.method === 'GET' && url.pathname === '/api/strategies') return json(res, 200, strategies.list());
    if (req.method === 'POST' && url.pathname === '/api/strategies') {
      const body = await readBody(req); return json(res, 201, addStrategy(body));
    }
    if (req.method === 'GET' && url.pathname === '/api/providers') return json(res, 200, [
      { id: 'mock', name: 'داده آزمایشی', state: 'ready', description: 'داده مصنوعی برای ارزیابی موتور و محیط توسعه' },
      { id: 'binance', name: 'Binance', state: 'public', description: 'داده عمومی بازار؛ سفارش واقعی در این نسخه پشتیبانی نمی‌شود' },
      { id: 'soft4fx', name: 'Soft4FX / CSV', state: 'planned', description: 'آداپتر فایل CSV؛ اتصال مستقیم در نسخه بعدی' }
    ]);
    if (req.method === 'POST' && url.pathname === '/api/backtest') {
      const body = await readBody(req);
      const strategy = getStrategy(body.strategyId);
      const provider = createProvider(body.provider || 'mock');
      const candles = await provider.getCandles({ symbol: body.symbol, timeframe: body.timeframe || '1h', limit: 700 });
      const result = runBacktest({ strategy, candles, risk: body.risk, initialBalance: body.balance });
      if (result.summary.trades) learningStore.record({ strategyId: strategy.id, returnPct: result.summary.returnPct, drawdownPct: result.summary.maxDrawdownPct, winRatePct: result.summary.winRatePct, trades: result.summary.trades });
      return json(res, 200, result);
    }
    if (req.method === 'POST' && url.pathname === '/api/paper/start') {
      const body = await readBody(req);
      const strategy = getStrategy(body.strategyId);
      if (!body.confirmed) return json(res, 400, { error: 'برای شروع، حالت شبیه‌سازی را تأیید کنید.' });
      paperSession = { id: `paper-${Date.now()}`, strategyId: strategy.id, strategyName: strategy.name, symbol: body.symbol || 'BTCUSDT', risk: Number(body.risk || 1), startedAt: new Date().toISOString(), status: 'running', balance: Number(body.balance || 10000) };
      return json(res, 200, paperSession);
    }
    if (req.method === 'POST' && url.pathname === '/api/paper/stop') { paperSession = null; return json(res, 200, { stopped: true }); }
    if (req.method === 'GET' && url.pathname === '/api/paper/status') return json(res, 200, paperSession || { status: 'idle' });
    if (req.method === 'GET') return serveStatic(req, res, url.pathname);
    return json(res, 404, { error: 'مسیر پیدا نشد' });
  } catch (error) { return json(res, error.status || 400, { error: error.message || 'خطای غیرمنتظره' }); }
});
server.listen(PORT, '0.0.0.0', () => console.log(`Trading Lab listening on 0.0.0.0:${PORT}`));
