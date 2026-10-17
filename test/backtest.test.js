const test = require('node:test');
const assert = require('node:assert/strict');
const { runBacktest, ema, rsi, signalsFor } = require('../src/backtest');
const { MockProvider } = require('../src/providers');

test('EMA returns values only after its warm-up period', () => {
  assert.deepEqual(ema([1, 2, 3, 4], 3).slice(0, 2), [null, null]);
  assert.equal(ema([1, 2, 3, 4], 3)[2], 2.25);
});
test('RSI computes a bounded oscillator', () => {
  const result = rsi(Array.from({ length: 30 }, (_, i) => i + 1), 14);
  assert.equal(result[14], 100);
  assert.ok(result.every(v => v === null || (v >= 0 && v <= 100)));
});
test('all built-in signal families produce correctly sized signal arrays', () => {
  const candles = Array.from({ length: 100 }, (_, i) => ({ open: 100 + Math.sin(i / 3), high: 105 + Math.sin(i / 3), low: 95 + Math.sin(i / 3), close: 100 + Math.sin(i / 3) }));
  for (const strategy of [
    { type: 'ema-cross', params: { fast: 4, slow: 12 } },
    { type: 'rsi-reversal', params: { period: 14, oversold: 30, overbought: 70 } },
    { type: 'breakout', params: { lookback: 10 } }
  ]) assert.equal(signalsFor(strategy, candles).length, candles.length);
});
test('synthetic provider is deterministic and returns OHLC candles', async () => {
  const provider = new MockProvider();
  const a = await provider.getCandles({ symbol: 'BTCUSDT', limit: 100 });
  const b = await provider.getCandles({ symbol: 'BTCUSDT', limit: 100 });
  assert.deepEqual(a.map(c => [c.open, c.high, c.low, c.close]), b.map(c => [c.open, c.high, c.low, c.close]));
  assert.ok(a.every(c => c.high >= Math.max(c.open, c.close) && c.low <= Math.min(c.open, c.close)));
});
test('backtest returns finite metrics and respects non-negative account balance', async () => {
  const candles = await new MockProvider().getCandles({ symbol: 'BTCUSDT', limit: 350 });
  const result = runBacktest({ strategy: { type: 'ema-cross', params: { fast: 8, slow: 22 } }, candles, risk: 1, initialBalance: 10000 });
  assert.ok(Number.isFinite(result.summary.returnPct));
  assert.ok(result.summary.finalBalance > 0);
  assert.equal(result.summary.riskPct, 1);
  assert.ok(result.equity.length > 0);
});
test('backtest rejects insufficient candle history', () => {
  assert.throws(() => runBacktest({ strategy: { type: 'ema-cross', params: { fast: 2, slow: 3 } }, candles: [], risk: 1 }), /حداقل ۵۰ کندل/);
});
