function ema(values, period) {
  const out = Array(values.length).fill(null); const k = 2 / (period + 1); let value = null;
  values.forEach((v, i) => { value = value === null ? v : v * k + value * (1 - k); if (i >= period - 1) out[i] = value; }); return out;
}
function rsi(values, period) {
  const out = Array(values.length).fill(null); let gain = 0, loss = 0;
  for (let i = 1; i <= period; i++) { const d = values[i] - values[i - 1]; gain += Math.max(0, d); loss += Math.max(0, -d); }
  let avgGain = gain / period, avgLoss = loss / period;
  out[period] = avgLoss === 0 ? 100 : 100 - 100 / (1 + avgGain / avgLoss);
  for (let i = period + 1; i < values.length; i++) { const d = values[i] - values[i - 1]; avgGain = (avgGain * (period - 1) + Math.max(0, d)) / period; avgLoss = (avgLoss * (period - 1) + Math.max(0, -d)) / period; out[i] = avgLoss === 0 ? 100 : 100 - 100 / (1 + avgGain / avgLoss); }
  return out;
}
function signalsFor(strategy, candles) {
  const closes = candles.map(c => c.close), signals = Array(candles.length).fill(0), p = strategy.params;
  if (strategy.type === 'ema-cross') {
    const fast = ema(closes, p.fast), slow = ema(closes, p.slow);
    for (let i = 1; i < closes.length; i++) if (fast[i] && slow[i] && fast[i - 1] && slow[i - 1]) {
      if (fast[i] > slow[i] && fast[i - 1] <= slow[i - 1]) signals[i] = 1;
      else if (fast[i] < slow[i] && fast[i - 1] >= slow[i - 1]) signals[i] = -1;
    }
  } else if (strategy.type === 'rsi-reversal') {
    const values = rsi(closes, p.period);
    for (let i = 1; i < closes.length; i++) if (values[i] != null && values[i - 1] != null) {
      if (values[i] > p.oversold && values[i - 1] <= p.oversold) signals[i] = 1;
      else if (values[i] < p.overbought && values[i - 1] >= p.overbought) signals[i] = -1;
    }
  } else if (strategy.type === 'breakout') {
    const n = p.lookback;
    for (let i = n; i < candles.length; i++) {
      let high = -Infinity, low = Infinity;
      for (let j = i - n; j < i; j++) { high = Math.max(high, candles[j].high); low = Math.min(low, candles[j].low); }
      if (candles[i].close > high) signals[i] = 1; else if (candles[i].close < low) signals[i] = -1;
    }
  }
  return signals;
}
function runBacktest({ strategy, candles, risk = 1, initialBalance = 10000 }) {
  if (!Array.isArray(candles) || candles.length < 50) throw new Error('برای بک‌تست حداقل ۵۰ کندل لازم است.');
  const starting = Math.max(100, Number(initialBalance) || 10000);
  const riskPct = Math.min(5, Math.max(0.1, Number(risk) || 1)) / 100;
  const signals = signalsFor(strategy, candles); let balance = starting, peak = starting, maxDrawdown = 0, position = null;
  const trades = []; const equity = [];
  const feeRate = 0.0005;
  function closePosition(i, price, reason) {
    const gross = position.side === 1 ? (price - position.entry) / position.entry * position.notional : (position.entry - price) / position.entry * position.notional;
    const net = gross - position.notional * feeRate * 2;
    balance += net;
    trades.push({ entryTime: candles[position.index].time, exitTime: candles[i].time, side: position.side === 1 ? 'خرید' : 'فروش', entry: position.entry, exit: price, pnl: net, pnlPct: net / position.notional * 100, stopLossPct: position.stopPct, takeProfitPct: position.targetPct, reason }); position = null;
  }
  function stopDistanceAt(index) {
    const start = Math.max(1, index - 13); let total = 0, count = 0;
    for (let j = start; j <= index; j++) { const prev = candles[j - 1].close; total += Math.max(candles[j].high - candles[j].low, Math.abs(candles[j].high - prev), Math.abs(candles[j].low - prev)); count++; }
    return Math.min(.05, Math.max(.005, (total / Math.max(1, count)) / candles[index].close * 1.5));
  }
  for (let i = 1; i < candles.length; i++) {
    let stoppedOnThisBar = false;
    // Protective stop derives from recent ATR, clamped to 0.5–5%; target uses 2R.
    // If both stop and target fall inside the same OHLC candle, the stop is assumed first.
    if (position) {
      const stopPrice = position.entry * (1 - position.side * position.stopPct);
      const targetPrice = position.entry * (1 + position.side * position.targetPct);
      const stopped = position.side === 1 ? candles[i].low <= stopPrice : candles[i].high >= stopPrice;
      const targeted = position.side === 1 ? candles[i].high >= targetPrice : candles[i].low <= targetPrice;
      if (stopped) { closePosition(i, stopPrice, `حد ضرر پویا ${ (position.stopPct * 100).toFixed(2) }٪`); stoppedOnThisBar = true; }
      else if (targeted) { closePosition(i, targetPrice, `حد سود ${ (position.targetPct * 100).toFixed(2) }٪`); stoppedOnThisBar = true; }
    }
    if (signals[i] && position && signals[i] !== position.side) closePosition(i, candles[i].close, 'تغییر سیگنال');
    if (signals[i] && !position && !stoppedOnThisBar) {
      const stopPct = stopDistanceAt(i), targetPct = stopPct * 2;
      const notional = Math.min(balance * 0.95, balance * riskPct / stopPct);
      position = { side: signals[i], entry: candles[i].close, index: i, notional, stopPct, targetPct };
    }
    const unrealized = position ? (position.side === 1 ? candles[i].close - position.entry : position.entry - candles[i].close) / position.entry * position.notional : 0;
    const value = balance + unrealized; equity.push({ time: candles[i].time, value }); peak = Math.max(peak, value); maxDrawdown = Math.max(maxDrawdown, (peak - value) / peak);
  }
  if (position) closePosition(candles.length - 1, candles.at(-1).close, 'پایان بازه');
  const wins = trades.filter(t => t.pnl > 0); const losses = trades.filter(t => t.pnl <= 0);
  const grossWin = wins.reduce((s, t) => s + t.pnl, 0), grossLoss = Math.abs(losses.reduce((s, t) => s + t.pnl, 0));
  const net = balance - starting;
  return { summary: { initialBalance: starting, finalBalance: balance, netProfit: net, returnPct: net / starting * 100, maxDrawdownPct: maxDrawdown * 100, winRatePct: trades.length ? wins.length / trades.length * 100 : 0, profitFactor: grossLoss ? grossWin / grossLoss : grossWin ? null : 0, trades: trades.length, wins: wins.length, losses: losses.length, riskPct: riskPct * 100 }, trades: trades.slice(-30).reverse(), equity, candles: candles.slice(-180).map(c => ({ time: c.time, close: c.close })), disclaimer: 'نتایج گذشته یا داده مصنوعی تضمین‌کننده سود آینده نیست. کارمزد ساده‌سازی شده و لغزش قیمت/تأخیر/تأمین مالی لحاظ نشده است.' };
}
module.exports = { runBacktest, signalsFor, ema, rsi };
