const { runBacktest } = require('./backtest');
function parameterGrid(strategy) {
  if (strategy.type === 'ema-cross') return [6, 8, 10, 12, 16].flatMap(fast => [20, 26, 34, 42].filter(slow => fast < slow).map(slow => ({ fast, slow })));
  if (strategy.type === 'rsi-reversal') return [10, 14, 21].flatMap(period => [25, 30, 35].map(oversold => ({ period, oversold, overbought: 100 - oversold })));
  if (strategy.type === 'breakout') return [8, 12, 20, 30, 40].map(lookback => ({ lookback }));
  return [strategy.params];
}
function optimize({ strategy, candles, risk = 1, initialBalance = 10000 }) {
  const candidates = parameterGrid(strategy).map(params => {
    const result = runBacktest({ strategy: { ...strategy, params }, candles, risk, initialBalance });
    const s = result.summary;
    const score = s.returnPct - s.maxDrawdownPct * .75 + Math.min(s.trades, 20) * .03;
    return { params, score, returnPct: s.returnPct, maxDrawdownPct: s.maxDrawdownPct, winRatePct: s.winRatePct, trades: s.trades, profitFactor: s.profitFactor };
  }).sort((a, b) => b.score - a.score);
  return { best: candidates[0], candidates: candidates.slice(0, 8), tested: candidates.length, objective: 'بازده خالص - ۰.۷۵ × افت سرمایه + امتیاز جزئی تعداد معاملات', warning: 'بهینه‌سازی روی همین بازه مستعد بیش‌برازش است. پارامتر برتر را پیش از استفاده، روی داده خارج از نمونه و paper trading ارزیابی کنید؛ این فرایند خودکار مدل یادگیری ماشین نیست.' };
}
module.exports = { optimize };
