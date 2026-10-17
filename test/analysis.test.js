const test = require('node:test');
const assert = require('node:assert/strict');
const { analyze, marketClock, rsi } = require('../src/analysis');
const { DemoProvider, scoreHeadline } = require('../src/providers');

test('demo mode returns indicators, risk levels, position size, and no invented fundamentals/news', async () => {
  const symbol = 'AAPL'; const candles = await new DemoProvider().getCandles({ symbol, limit: 260 });
  const quote = await new DemoProvider().getQuote({ candles });
  const result = analyze({ symbol, candles, quote, fundamentals: await new DemoProvider().getFundamentals(), news: [], riskPct: 1, balance: 10000 });
  assert.equal(result.symbol, symbol);
  assert.ok(result.technical.atr > 0);
  assert.ok(result.technical.rsi >= 0 && result.technical.rsi <= 100);
  assert.ok(result.plan.stopPct >= .5 && result.plan.stopPct <= 7);
  assert.equal(result.plan.targetPct, result.plan.stopPct * 2);
  assert.ok(result.plan.shares >= 0);
  assert.equal(result.fundamentals.available, false);
  assert.equal(result.news.available, false);
  assert.equal(result.pressure.buySharePct + result.pressure.sellSharePct, 100);
});
test('demo candles respect selected chart interval', async () => {
  const candles = await new DemoProvider().getCandles({ symbol: 'MSFT', timeframe: '1h', limit: 60 });
  assert.equal(candles[1].time - candles[0].time, 60 * 60 * 1000);
});
test('US market hours use New York timezone and regular-session bounds', () => {
  const open = marketClock(new Date('2026-10-19T14:00:00.000Z'));
  const closed = marketClock(new Date('2026-10-19T21:00:00.000Z'));
  assert.equal(open.regularSession, true); // 10:00 EDT Monday
  assert.equal(closed.regularSession, false); // 17:00 EDT Monday
  assert.equal(open.timeZone, 'America/New_York');
});
test('headline sentiment is explicitly keyword based', () => {
  assert.match(scoreHeadline('Company beats earnings estimates'), /مثبت/);
  assert.match(scoreHeadline('Company reports quarterly loss'), /منفی/);
  assert.match(scoreHeadline('Quarterly company update'), /خنثی/);
});
test('RSI calculation stays bounded for falling prices', () => {
  const out = rsi(Array.from({ length: 40 }, (_, i) => 100 - i), 14);
  assert.equal(out, 0);
});
