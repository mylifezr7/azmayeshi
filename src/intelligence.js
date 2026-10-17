const { signalsFor, ema } = require('./backtest');
function analyzeMarket({ candles, strategies, requestedRisk = 1, learning = [] }) {
  const closes = candles.map(x => x.close), latest = candles.at(-1), n = candles.length;
  if (n < 50) throw new Error('برای تحلیل حداقل ۵۰ کندل لازم است.');
  const fast = ema(closes, 12), slow = ema(closes, 26);
  const trend = fast.at(-1) > slow.at(-1) ? 'صعودی' : 'نزولی';
  const momentum5 = (closes.at(-1) / closes.at(-6) - 1) * 100;
  const momentum20 = (closes.at(-1) / closes.at(-21) - 1) * 100;
  const ranges = candles.slice(-15).map((c, idx) => {
    const prev = candles[Math.max(0, n - 15 + idx - 1)].close;
    return Math.max(c.high - c.low, Math.abs(c.high - prev), Math.abs(c.low - prev));
  });
  const atr = ranges.reduce((s, x) => s + x, 0) / ranges.length;
  const atrPct = atr / latest.close * 100;
  const stopPct = Math.min(5, Math.max(.5, atrPct * 1.5));
  const targetPct = stopPct * 2;
  const volumes = candles.slice(-21, -1).map(c => c.volume || 0);
  const avgVolume = volumes.reduce((s, x) => s + x, 0) / (volumes.length || 1);
  const currentVolume = latest.volume || 0;
  const volumeRatio = avgVolume ? currentVolume / avgVolume : 1;
  const spreadProxyPct = ((latest.high - latest.low) / latest.close) * 100;
  const liquidityScore = Math.max(0, Math.min(100, 85 - spreadProxyPct * 60 + Math.min(15, volumeRatio * 5)));
  const signals = strategies.map(strategy => {
    const all = signalsFor(strategy, candles);
    let found = 0, age = Infinity;
    for (let i = all.length - 1; i >= Math.max(0, all.length - 8); i--) if (all[i]) { found = all[i]; age = all.length - 1 - i; break; }
    return { strategyId: strategy.id, name: strategy.name, signal: found === 1 ? 'خرید' : found === -1 ? 'فروش' : 'خنثی', ageBars: Number.isFinite(age) ? age : null };
  });
  const active = signals.filter(s => s.signal !== 'خنثی');
  const directionBalance = active.reduce((sum, s) => sum + (s.signal === 'خرید' ? 1 : -1), 0);
  const agreement = active.length ? Math.round(Math.abs(directionBalance) / active.length * 100) : 0;
  const direction = directionBalance > 0 ? 'خرید' : directionBalance < 0 ? 'فروش' : 'بدون اجماع';
  const volatilityPenalty = Math.min(30, Math.max(0, (atrPct - 1) * 10));
  const confidence = Math.max(0, Math.min(100, Math.round(agreement * .55 + Math.min(100, Math.abs(momentum20) * 10) * .25 + (trend === direction ? 20 : 0) - volatilityPenalty)));
  const requested = Math.max(.1, Math.min(5, Number(requestedRisk) || 1));
  const riskAdjustedPct = Math.min(requested, liquidityScore < 35 || atrPct > 4 ? requested * .5 : requested);
  const learned = learning.find(x => x.strategyId === (active[0] && active[0].strategyId));
  return {
    symbolPrice: latest.close, trend, momentum5Pct: momentum5, momentum20Pct: momentum20, atrPct,
    liquidity: { score: Math.round(liquidityScore), label: liquidityScore >= 70 ? 'نسبتاً مناسب' : liquidityScore >= 40 ? 'متوسط' : 'ضعیف', avgVolume, latestVolume: currentVolume, candleRangeProxyPct: spreadProxyPct, note: 'برآورد از حجم و دامنه کندل؛ داده دفتر سفارش/جریان سفارش و نهنگ‌ها متصل نیست.' },
    sentiment: { score: 50, label: 'نامشخص', source: 'غیرفعال', note: 'خوراک خبر یا شبکه اجتماعی وصل نیست؛ این امتیاز خنثی/پیش‌فرض است و تحلیل احساسات واقعی محسوب نمی‌شود.' },
    signals, decision: { direction, confidencePct: confidence, agreementPct: agreement, layers: [{ name: 'روند EMA', value: trend }, { name: 'مومنتوم ۲۰ کندل', value: `${momentum20 >= 0 ? '+' : ''}${momentum20.toFixed(2)}٪` }, { name: 'اجماع استراتژی‌ها', value: `${active.length}/${strategies.length}` }, { name: 'کنترل نوسان/نقدشوندگی', value: riskAdjustedPct < requested ? 'ریسک کاهش یافت' : 'عادی' }], note: 'امتیاز اکتشافی و قاعده‌محور است، نه سیگنال قطعی یا توصیه معامله.' },
    risk: { requestedRiskPct: requested, adjustedRiskPct: riskAdjustedPct, stopLossPct: stopPct, takeProfitPct: targetPct, riskReward: '۱:۲', sizingBasis: 'حد ضرر پویا بر پایه ATR تقریبی؛ شکاف بازار می‌تواند زیان را افزایش دهد.' },
    learning: learned ? { ...learned, status: 'نمونه‌های ثبت‌شده؛ بدون آموزش مدل یادگیری ماشین' } : { samples: 0, status: 'هنوز نتیجه‌ای برای این ستاپ ثبت نشده است.' }
  };
}
module.exports = { analyzeMarket };
