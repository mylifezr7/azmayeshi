const { scoreHeadline } = require('./providers');
function ema(values, period) {
  const out = Array(values.length).fill(null); const k = 2 / (period + 1); let v = null;
  values.forEach((x, i) => { v = v === null ? x : x * k + v * (1 - k); if (i >= period - 1) out[i] = v; }); return out;
}
function rsi(values, period = 14) {
  if (values.length <= period) return null;
  let up = 0, down = 0;
  for (let i = 1; i <= period; i++) { const d = values[i] - values[i - 1]; up += Math.max(0, d); down += Math.max(0, -d); }
  let avgUp = up / period, avgDown = down / period;
  for (let i = period + 1; i < values.length; i++) { const d = values[i] - values[i - 1]; avgUp = (avgUp * (period - 1) + Math.max(0, d)) / period; avgDown = (avgDown * (period - 1) + Math.max(0, -d)) / period; }
  return avgDown === 0 ? 100 : 100 - 100 / (1 + avgUp / avgDown);
}
function marketClock(now = new Date()) {
  const fmt = new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', weekday: 'short', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' });
  const parts = Object.fromEntries(fmt.formatToParts(now).map(p => [p.type, p.value]));
  const minute = Number(parts.hour) * 60 + Number(parts.minute);
  const weekday = !['Sat', 'Sun'].includes(parts.weekday);
  const regular = weekday && minute >= 570 && minute < 960;
  let phase = regular ? 'بازار اصلی باز است' : weekday && minute >= 540 && minute < 570 ? 'پیش‌گشایش' : weekday && minute >= 960 && minute < 1200 ? 'پس از بازار' : 'بازار اصلی بسته است';
  let next = null;
  if (!regular) {
    for (let delta = 0; delta < 9 * 24 * 60; delta += 15) {
      const test = new Date(now.getTime() + delta * 60000);
      const p = Object.fromEntries(fmt.formatToParts(test).map(x => [x.type, x.value]));
      if (!['Sat', 'Sun'].includes(p.weekday) && Number(p.hour) * 60 + Number(p.minute) >= 570 && Number(p.hour) * 60 + Number(p.minute) < 585) { next = test.toISOString(); break; }
    }
  }
  return { phase, regularSession: regular, timeZone: 'America/New_York', localTime: `${parts.weekday} ${parts.hour}:${parts.minute}`, nextOpenUtc: next, note: 'برنامه بر اساس روزهای کاری و ساعات ۹:۳۰ تا ۱۶:۰۰ نیویورک است؛ تعطیلات رسمی بورس در این برآورد لحاظ نمی‌شود.' };
}
function analyze({ symbol, candles, quote, fundamentals, news = [], riskPct = 1, balance = 10000, timeframe = '15m' }) {
  if (!Array.isArray(candles) || candles.length < 50) throw new Error('برای تحلیل تکنیکال حداقل ۵۰ کندل لازم است.');
  const closes = candles.map(c => Number(c.close));
  const lastCandle = candles.at(-1), price = Number(quote?.price || lastCandle.close);
  const fast = ema(closes, 12), slow = ema(closes, 26);
  const macd = closes.map((_, i) => fast[i] != null && slow[i] != null ? fast[i] - slow[i] : null);
  const macdSeries = ema(macd.map(v => v ?? 0), 9);
  const macdHist = (macd.at(-1) || 0) - (macdSeries.at(-1) || 0);
  const rsiValue = rsi(closes, 14);
  const trueRanges = [];
  for (let i = Math.max(1, candles.length - 14); i < candles.length; i++) { const prev = closes[i - 1]; trueRanges.push(Math.max(candles[i].high - candles[i].low, Math.abs(candles[i].high - prev), Math.abs(candles[i].low - prev))); }
  const atr = trueRanges.reduce((s, x) => s + x, 0) / trueRanges.length;
  const window = candles.slice(-21, -1);
  const resistance = Math.max(...window.map(c => c.high)), support = Math.min(...window.map(c => c.low));
  let buyVol = 0, sellVol = 0;
  for (const candle of candles.slice(-20)) {
    const range = Math.max(.00000001, candle.high - candle.low);
    const buyShare = Math.max(0, Math.min(1, (candle.close - candle.low) / range));
    buyVol += (candle.volume || 0) * buyShare; sellVol += (candle.volume || 0) * (1 - buyShare);
  }
  const totalFlow = buyVol + sellVol;
  const pressurePct = totalFlow ? (buyVol - sellVol) / totalFlow * 100 : 0;
  const trendUp = fast.at(-1) > slow.at(-1);
  const technicalParts = [
    { label: 'روند EMA ۱۲/۲۶', value: trendUp ? 'صعودی' : 'نزولی', score: trendUp ? 1 : -1 },
    { label: 'MACD', value: macdHist >= 0 ? 'مثبت' : 'منفی', score: macdHist >= 0 ? 1 : -1 },
    { label: 'RSI ۱۴', value: rsiValue == null ? 'ناموجود' : rsiValue.toFixed(1), score: rsiValue != null && rsiValue < 35 ? 1 : rsiValue != null && rsiValue > 70 ? -1 : 0 },
    { label: 'شکست محدوده ۲۰ کندل', value: price > resistance ? 'شکست سقف' : price < support ? 'شکست کف' : 'داخل محدوده', score: price > resistance ? 1 : price < support ? -1 : 0 }
  ];
  const technicalScore = technicalParts.reduce((s, x) => s + x.score, 0);
  const flowScore = pressurePct > 12 ? 1 : pressurePct < -12 ? -1 : 0;
  const metric = fundamentals || {};
  const fundamentalParts = [];
  if (Number.isFinite(metric.revenueGrowth)) fundamentalParts.push({ label: 'رشد درآمد سالانه', value: `${metric.revenueGrowth.toFixed(1)}٪`, score: metric.revenueGrowth > 5 ? 1 : metric.revenueGrowth < -5 ? -1 : 0 });
  if (Number.isFinite(metric.roe)) fundamentalParts.push({ label: 'بازده حقوق صاحبان سهام', value: `${metric.roe.toFixed(1)}٪`, score: metric.roe > 10 ? 1 : metric.roe < 0 ? -1 : 0 });
  if (Number.isFinite(metric.debtToEquity)) fundamentalParts.push({ label: 'بدهی به حقوق صاحبان سهام', value: `${metric.debtToEquity.toFixed(1)}`, score: metric.debtToEquity > 200 ? -1 : 0 });
  if (Number.isFinite(metric.pe)) fundamentalParts.push({ label: 'P/E دوازده‌ماهه', value: `${metric.pe.toFixed(1)}`, score: metric.pe < 0 || metric.pe > 80 ? -1 : 0 });
  const fundamentalScore = fundamentalParts.reduce((s, x) => s + x.score, 0);
  const headlineRows = news.map(item => ({ ...item, sentiment: item.sentiment || scoreHeadline(item.headline || '') }));
  const newsScore = headlineRows.reduce((s, item) => s + (item.sentiment.startsWith('مثبت') ? 1 : item.sentiment.startsWith('منفی') ? -1 : 0), 0);
  const score = technicalScore + flowScore + fundamentalScore + Math.max(-1, Math.min(1, newsScore));
  const direction = score >= 3 ? 'تمایل صعودی' : score <= -3 ? 'تمایل نزولی' : 'خنثی / انتظار';
  const confidence = Math.min(95, Math.round(48 + Math.abs(score) * 8));
  const stopPct = Math.min(7, Math.max(.5, atr / price * 1.5 * 100));
  const stopDistance = price * stopPct / 100;
  const stop = score < 0 ? price + stopDistance : price - stopDistance;
  const target = score < 0 ? price - stopDistance * 2 : price + stopDistance * 2;
  const safeRisk = Math.min(5, Math.max(.1, Number(riskPct) || 1));
  const capital = Math.max(100, Number(balance) || 10000);
  const riskAmount = capital * safeRisk / 100;
  const sharesByRisk = Math.floor(riskAmount / Math.max(.000001, stopDistance));
  const maxAffordable = Math.floor(capital * .95 / price);
  const suggestedShares = score > 0 ? Math.max(0, Math.min(sharesByRisk, maxAffordable)) : 0;
  const riskDollars = suggestedShares * stopDistance;
  const hasFundamentalMetrics = ['marketCap', 'pe', 'eps', 'revenueGrowth', 'debtToEquity', 'roe'].some(key => metric[key] != null && Number.isFinite(Number(metric[key])));
  return {
    symbol: symbol.toUpperCase(), timeframe, quote: { price, previousClose: quote?.previousClose ?? null, changePct: quote?.changePct ?? null, source: quote?.source || 'نامشخص', updatedAt: quote?.updatedAt || lastCandle.time },
    session: marketClock(),
    technical: { trend: trendUp ? 'صعودی' : 'نزولی', rsi: rsiValue, macdHistogram: macdHist, atr, atrPct: atr / price * 100, support, resistance, score: technicalScore, factors: technicalParts },
    pressure: { buyVolume: buyVol, sellVolume: sellVol, buySharePct: totalFlow ? buyVol / totalFlow * 100 : 50, sellSharePct: totalFlow ? sellVol / totalFlow * 100 : 50, pressurePct, label: pressurePct > 12 ? 'غلبهٔ فشار خرید' : pressurePct < -12 ? 'غلبهٔ فشار فروش' : 'فشار تقریباً متعادل', note: 'تخمین از حجم و محل بسته‌شدن کندل؛ دادهٔ واقعی تیک/دفتر سفارش نیست.' },
    fundamentals: metric?.source ? { ...metric, score: fundamentalScore, factors: fundamentalParts, available: hasFundamentalMetrics } : { available: false, factors: [], score: 0 },
    news: { items: headlineRows.slice(0, 8), score: newsScore, label: headlineRows.length ? 'برچسب واژه‌محور تیترها' : 'خبر در دسترس نیست', available: headlineRows.length > 0, note: 'برچسب مثبت/منفی با کلیدواژه‌های سادهٔ انگلیسی است و درک زمینه/فارسی ندارد.' },
    decision: { direction, score, confidencePct: confidence, note: 'امتیاز قاعده‌محور و بدون مدل هوش مصنوعی آموزش‌دیده است؛ قدرت اکتشافی بوده و احتمال موفقیت یا توصیه سرمایه‌گذاری محسوب نمی‌شود.' },
    plan: { referenceEntry: price, stopLoss: stop, takeProfit: target, stopPct, targetPct: stopPct * 2, riskReward: '۱:۲', riskPct: safeRisk, riskAmount, shares: suggestedShares, estimatedPositionValue: suggestedShares * price, note: score > 0 ? 'قیمت مرجع/ناحیهٔ بررسی است؛ سفارش یا ورود خودکار انجام نمی‌شود.' : 'سیگنال ورود قطعی نیست؛ در حالت خنثی/نزولی از ورود خرید اجتناب کنید یا برنامه را فقط برای سناریوی مشروط ببینید.' },
    dataNotice: 'دادهٔ دمو مصنوعی است. دادهٔ بنیادی/خبر فقط با اتصال Finnhub و تعیین FINNHUB_API_KEY می‌آید.',
    modelNotice: 'تحلیل بر پایه اندیکاتور و قواعد امتیازدهی است؛ مدل هوش مصنوعی آموزش‌دیده در این نسخه وجود ندارد.'
  };
}
module.exports = { analyze, marketClock, ema, rsi };
