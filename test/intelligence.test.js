const test = require('node:test');
const assert = require('node:assert/strict');
const { MockProvider } = require('../src/providers');
const { analyzeMarket } = require('../src/intelligence');
const { optimize } = require('../src/optimizer');
const { strategies } = require('../src/strategies');

test('market analysis includes multiple layers and explicit non-connected source labels', async () => {
  const candles = await new MockProvider().getCandles({ symbol: 'ETHUSDT', limit: 240 });
  const result = analyzeMarket({ candles, strategies: strategies.list(), requestedRisk: 2 });
  assert.ok(result.decision.confidencePct >= 0 && result.decision.confidencePct <= 100);
  assert.equal(result.signals.length, 3);
  assert.ok(result.risk.stopLossPct >= .5 && result.risk.stopLossPct <= 5);
  assert.equal(result.risk.takeProfitPct, result.risk.stopLossPct * 2);
  assert.match(result.liquidity.note, /دفتر سفارش/);
  assert.equal(result.sentiment.source, 'غیرفعال');
});
test('parameter search returns ranked candidates and warns against in-sample overfit', async () => {
  const candles = await new MockProvider().getCandles({ symbol: 'BTCUSDT', limit: 120 });
  const result = optimize({ strategy: strategies.list()[0], candles, risk: 1 });
  assert.equal(result.tested, 20);
  assert.ok(result.candidates.length > 0);
  assert.ok(Number.isFinite(result.best.score));
  assert.match(result.warning, /بیش‌برازش/);
});
