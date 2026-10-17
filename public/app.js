const $ = selector => document.querySelector(selector);
const fmt = (value, digits = 2) => value == null || !Number.isFinite(Number(value)) ? '—' : new Intl.NumberFormat('fa-IR', { maximumFractionDigits: digits }).format(Number(value));
const money = (value, digits = 2) => value == null || !Number.isFinite(Number(value)) ? '—' : '$' + new Intl.NumberFormat('en-US', { maximumFractionDigits: digits }).format(Number(value));
let toastTimer;
function toast(message) { const el = $('#toast'); el.textContent = message; el.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('show'), 4200); }
async function api(url) { const response = await fetch(url); const data = await response.json().catch(() => ({})); if (!response.ok) throw new Error(data.error || 'خطا در دریافت داده'); return data; }
function safe(value) { return String(value ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c])); }
function fmtPrice(price) { if (price == null) return '—'; const decimals = Number(price) < 1 ? 4 : 2; return '$' + new Intl.NumberFormat('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(price); }
function refreshHeaderSession(session) {
  $('#session-label').textContent = session.phase;
  $('#header-session').textContent = session.regularSession ? 'سشن اصلی باز است' : 'سشن اصلی بسته است';
  $('#banner-session').textContent = session.phase;
  $('#banner-session-note').textContent = session.regularSession ? 'پنجرهٔ اصلی معامله باز است: روز کاری، ۹:۳۰ تا ۱۶:۰۰ نیویورک.' : `بازگشایی بعدی (تقریبی): ${session.nextOpenUtc ? new Intl.DateTimeFormat('fa-IR', { timeZone: 'America/New_York', weekday: 'long', hour: '2-digit', minute: '2-digit' }).format(new Date(session.nextOpenUtc)) + ' به وقت نیویورک' : 'نامشخص'} · تعطیلات رسمی لحاظ نشده‌اند.`;
}
function renderFundamentals(fundamentals, status) {
  const available = Boolean(fundamentals && fundamentals.available);
  $('#fundamental-status').textContent = available ? 'داده دریافت شد' : 'داده در دسترس نیست';
  $('#fundamental-status').className = `state ${available ? 'ready' : 'planned'}`;
  const fields = [
    ['ارزش بازار', fundamentals?.marketCap != null ? `$${fmt(fundamentals.marketCap / 1000, 1)} B` : '—'],
    ['P/E دوازده‌ماهه', fmt(fundamentals?.pe)], ['EPS', fundamentals?.eps != null ? money(fundamentals.eps) : '—'],
    ['رشد درآمد سالانه', fundamentals?.revenueGrowth != null ? `${fmt(fundamentals.revenueGrowth)}٪` : '—'],
    ['ROE', fundamentals?.roe != null ? `${fmt(fundamentals.roe)}٪` : '—'], ['بدهی / حقوق صاحبان سهام', fmt(fundamentals?.debtToEquity)]
  ];
  $('#fundamental-cards').innerHTML = fields.map(([label, value]) => `<div class="fundamental-metric"><small>${label}</small><b>${safe(value)}</b></div>`).join('');
  $('#fundamental-foot').textContent = available ? `منبع: ${fundamentals.source || 'Finnhub'} · ارز ${fundamentals.currency || 'USD'} · این شاخص‌ها باید با گزارش رسمی شرکت تطبیق داده شوند.` : status?.fundamentalsError || 'برای دریافت اطلاعات بنیادی زنده، Finnhub API Key را روی سرور محلی تنظیم کن. در حالت دمو هیچ مقدار بنیادی ساختگی نمایش داده نمی‌شود.';
}
function renderNews(news, status) {
  const items = news?.items || [];
  $('#news-status').textContent = items.length ? `${fmt(items.length, 0)} تیتر` : 'داده نیست';
  $('#news-status').className = `state ${items.length ? 'public' : 'planned'}`;
  $('#news-list').innerHTML = items.length ? items.map(item => {
    const when = item.time ? new Intl.DateTimeFormat('fa-IR', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }).format(new Date(item.time)) : '';
    const tag = item.sentiment.startsWith('مثبت') ? 'news-positive' : item.sentiment.startsWith('منفی') ? 'news-negative' : 'news-neutral';
    const title = item.url && /^https:\/\//.test(item.url) ? `<a href="${safe(item.url)}" target="_blank" rel="noopener noreferrer">${safe(item.headline)}</a>` : safe(item.headline);
    return `<article class="news-item"><span class="news-tag ${tag}">${safe(item.sentiment)}</span><div><b>${title}</b><small>${safe(item.source || 'خبر')} · ${safe(when)}</small></div></article>`;
  }).join('') : `<div class="empty-news">${safe(status?.newsError || 'برای تیترهای تازه، Finnhub API Key را تنظیم کن. حالت دمو خبر واقعی یا ساختگی نشان نمی‌دهد.')}</div>`;
}
function renderAnalysis(data) {
  $('#connection-status').innerHTML = `<i></i> ${data.dataStatus?.provider === 'finnhub' ? 'منبع Finnhub' : 'حالت دمو'}`;
  $('#company-title').textContent = data.symbol;
  $('#company-subtitle').textContent = data.fundamentals?.name ? `${data.fundamentals.name} · ${data.fundamentals.exchange || 'US'}` : `بازه ${data.timeframe} · ${data.dataStatus?.provider === 'demo' ? 'کندل مصنوعی' : 'داده بازار'}`;
  $('#quote-source').textContent = `· ${data.quote.source || '—'}`;
  $('#quote-price').textContent = fmtPrice(data.quote.price);
  $('#quote-change').textContent = data.quote.changePct == null ? (data.quote.previousClose != null ? `بسته قبلی ${fmtPrice(data.quote.previousClose)}` : 'تغییر روزانه در دسترس نیست') : `${data.quote.changePct >= 0 ? '+' : ''}${fmt(data.quote.changePct)}٪ ${data.quote.changeLabel || 'تغییر قیمت'}`;
  $('#quote-change').className = data.quote.changePct == null ? '' : data.quote.changePct >= 0 ? 'positive' : 'negative';
  refreshHeaderSession(data.session);
  $('#data-badge').textContent = data.dataStatus?.provider === 'demo' ? 'دمو · مصنوعی' : 'دادهٔ بیرونی';
  $('#trend-value').textContent = data.technical.trend;
  $('#trend-icon').textContent = data.technical.trend === 'صعودی' ? '↗' : '↘';
  $('#trend-icon').className = `trend-icon ${data.technical.trend === 'صعودی' ? 'positive' : 'negative'}`;
  $('#rsi-pill').textContent = `RSI ${fmt(data.technical.rsi, 1)}`;
  $('#rsi-pill').className = `neutral-pill ${data.technical.rsi < 35 ? 'pill-positive' : data.technical.rsi > 70 ? 'pill-negative' : ''}`;
  const factors = [
    ['MACD histogram', data.technical.macdHistogram == null ? '—' : `${data.technical.macdHistogram >= 0 ? '+' : ''}${fmt(data.technical.macdHistogram, 3)}`],
    ['حمایت تقریبی · ۲۰ کندل', fmtPrice(data.technical.support)], ['مقاومت تقریبی · ۲۰ کندل', fmtPrice(data.technical.resistance)],
    ['ATR · دامنه واقعی', `${fmt(data.technical.atrPct)}٪`]
  ];
  $('#indicator-list').innerHTML = factors.map(([k,v]) => `<div><span>${safe(k)}</span><b>${safe(v)}</b></div>`).join('');
  $('#technical-score').textContent = `${data.technical.score > 0 ? '+' : ''}${fmt(data.technical.score, 0)} / ۴`;
  $('#pressure-label').textContent = data.pressure.label;
  $('#pressure-value').textContent = `${data.pressure.pressurePct >= 0 ? '+' : ''}${fmt(data.pressure.pressurePct)}٪ خالص تقریبی`;
  $('#buy-bar').style.width = `${Math.min(100, Math.max(0, data.pressure.buySharePct))}%`;
  $('#buy-share').textContent = `${fmt(data.pressure.buySharePct)}٪`;
  $('#sell-share').textContent = `${fmt(data.pressure.sellSharePct)}٪`;
  $('#pressure-note').textContent = data.pressure.note;
  $('#decision-label').textContent = data.decision.direction;
  $('#confidence-badge').textContent = `قدرت اکتشافی ${fmt(data.decision.confidencePct, 0)}٪`;
  $('#decision-meter-fill').style.width = `${data.decision.confidencePct}%`;
  $('#decision-meter-fill').style.background = data.decision.score > 0 ? '#62d9ad' : data.decision.score < 0 ? '#ff7e8b' : '#a99cff';
  $('#decision-note').textContent = data.decision.note;
  $('#combined-score').textContent = `${data.decision.score > 0 ? '+' : ''}${fmt(data.decision.score, 0)}`;
  $('#entry-price').textContent = fmtPrice(data.plan.referenceEntry);
  $('#entry-note').textContent = data.decision.score <= 0 ? 'سناریوی خرید فعال نیست؛ این قیمت صرفاً مرجع تحلیل است.' : data.session.regularSession ? 'سناریوی صعودی در سشن اصلی؛ پیش از اقدام مستقل تأیید کن.' : 'سشن اصلی بسته است؛ سطح مرجع است و قیمت اجرایی نیست.';
  $('#stop-price').textContent = fmtPrice(data.plan.stopLoss);
  $('#stop-percent').textContent = `${fmt(data.plan.stopPct)}٪ فاصله از قیمت مرجع`;
  $('#target-price').textContent = fmtPrice(data.plan.takeProfit);
  $('#target-percent').textContent = `۲R · ${fmt(data.plan.targetPct)}٪ فاصله`;
  $('#shares-count').textContent = data.plan.shares > 0 ? `${fmt(data.plan.shares, 0)} سهم · سناریوی صعودی` : 'ورود خرید پیشنهاد نمی‌شود';
  $('#risk-dollars').textContent = data.plan.shares > 0 ? `زیان هدف تقریبی ${money(data.plan.riskAmount)} · ارزش موقعیت ${money(data.plan.estimatedPositionValue)}` : 'در حالت خنثی/نزولی، حجم خرید محاسبه نمی‌شود.';
  $('#risk-warning').textContent = `${data.plan.note} حدضرر مبتنی بر ATR ${fmt(data.plan.stopPct)}٪ است؛ ${data.session.note}`;
  renderFundamentals(data.fundamentals, data.dataStatus);
  renderNews(data.news, data.dataStatus);
  $('#provider-note').textContent = data.dataStatus?.provider === 'demo' ? data.dataNotice : `${data.dataStatus?.fundamentals} فاندامنتال · ${data.dataStatus?.news} خبر · ${data.dataStatus?.quote} قیمت`;
  $('#overview-placeholder')?.remove();
}
async function analyzeSymbol() {
  const symbol = $('#symbol').value.trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9.\-]{0,11}$/.test(symbol)) { toast('نماد معتبر وارد کن؛ مانند AAPL یا BRK.B.'); return; }
  $('#symbol').value = symbol;
  const btn = $('#analyze-btn'); btn.disabled = true; btn.innerHTML = '<span>◌</span> در حال تحلیل…';
  try {
    const query = new URLSearchParams({ symbol, timeframe: $('#timeframe').value, provider: $('#provider').value, risk: $('#risk').value, balance: $('#balance').value });
    const result = await api(`/api/analysis?${query}`);
    renderAnalysis(result);
    if (result.dataStatus?.provider === 'finnhub' && (!result.fundamentals?.available || !result.news?.available)) toast(`تحلیل قیمت انجام شد؛ ${result.dataStatus.fundamentals} بنیادی / ${result.dataStatus.news} خبر.`);
  } catch (error) { toast(error.message); }
  finally { btn.disabled = false; btn.innerHTML = '<span>⌕</span> تحلیل این نماد <b>←</b>'; }
}
async function loadClock() { try { const clock = await api('/api/market-clock'); refreshHeaderSession(clock); } catch {} }
$('#analyze-btn').addEventListener('click', analyzeSymbol);
$('#symbol').addEventListener('keydown', e => { if (e.key === 'Enter') analyzeSymbol(); });
$('#provider').addEventListener('change', () => { const real = $('#provider').value === 'finnhub'; $('#provider-note').textContent = real ? 'Finnhub نیاز به API Key سروری دارد؛ کلید را داخل صفحه وارد نکن.' : 'دمو برای نمایش رابط است؛ قیمت و تحلیل آن واقعی نیست.'; });
$('#risk').addEventListener('input', e => { $('#risk-label').textContent = `${fmt(e.target.value, 1)}٪`; });
$('#timeframe').addEventListener('change', () => { const tf = $('#timeframe').value; if (tf === '1d') $('#provider-note').textContent = 'بازه روزانه برای تحلیل میان‌مدت است؛ ساعت بازار همچنان در America/New_York محاسبه می‌شود.'; });
$('#balance').addEventListener('change', analyzeSymbol);
$('#risk').addEventListener('change', analyzeSymbol);
loadClock(); analyzeSymbol();
setInterval(() => { if (!document.hidden) loadClock(); }, 60_000);
setInterval(() => { if (!document.hidden && $('#provider').value === 'finnhub') analyzeSymbol(); }, 5 * 60_000);
