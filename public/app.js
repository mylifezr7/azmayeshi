const $ = (s) => document.querySelector(s);
const fmt = (n, digits = 2) => new Intl.NumberFormat('fa-IR', { maximumFractionDigits: digits }).format(Number(n) || 0);
const fmtMoney = (n) => '$' + new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 }).format(Number(n) || 0);
let strategyData = [];
let toastTimer;
function toast(message) { const el = $('#toast'); el.textContent = message; el.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('show'), 3300); }
async function api(path, options = {}) {
  const response = await fetch(path, { headers: { 'Content-Type': 'application/json' }, ...options });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || 'خطا در ارتباط با سرور');
  return data;
}
function selectedStrategy() { return strategyData.find(s => s.id === $('#strategy-select').value); }
function renderStrategies(list) {
  strategyData = list;
  const select = $('#strategy-select'); const previous = select.value;
  select.innerHTML = list.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)} · ${escapeHtml(s.short)}</option>`).join('');
  if (list.some(s => s.id === previous)) select.value = previous;
  $('#nav-count').textContent = fmt(list.length, 0);
  $('#strategy-list').innerHTML = list.map((s, i) => `<div class="strategy-item"><span class="strategy-icon ${i % 3 === 1 ? 'mint' : i % 3 === 2 ? 'amber' : ''}">${i % 3 === 0 ? '⌁' : i % 3 === 1 ? '◉' : '↗'}</span><span class="strategy-copy"><b>${escapeHtml(s.name)}</b><small>${escapeHtml(s.description)}</small></span><span class="strategy-category">${escapeHtml(s.category)}</span><button class="strategy-use" data-use="${s.id}">انتخاب</button></div>`).join('');
  document.querySelectorAll('[data-use]').forEach(btn => btn.addEventListener('click', () => { select.value = btn.dataset.use; updateDescription(); document.querySelector('#backtest').scrollIntoView({ behavior: 'smooth', block: 'center' }); }));
  updateDescription();
}
function escapeHtml(value) { return String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function updateDescription() { const s = selectedStrategy(); $('#strategy-description').textContent = s ? `${s.description} · ${s.riskHint}` : ''; }
function chosenMode() { return document.querySelector('input[name="mode"]:checked').value; }
function updateMode() {
  const paper = chosenMode() === 'paper';
  document.querySelectorAll('.mode-option').forEach(el => el.classList.toggle('selected', el.querySelector('input').checked));
  $('#run-btn').innerHTML = paper ? '<span>◉</span> شروع اجرای آزمایشی <b>←</b>' : '<span>▶</span> اجرای بک‌تست <b>←</b>';
}
function drawChart(points) {
  const svg = $('#equity-chart');
  if (!points || points.length < 2) { svg.innerHTML = ''; return; }
  const vals = points.map(p => p.value); let min = Math.min(...vals), max = Math.max(...vals); if (max === min) { max += 1; min -= 1; }
  const pad = (max - min) * .12; min -= pad; max += pad;
  const coords = vals.map((v, i) => `${(i / (vals.length - 1) * 800).toFixed(2)},${(175 - (v - min) / (max - min) * 155).toFixed(2)}`);
  const firstX = 0, lastX = 800, y0 = 175 - (vals[0] - min) / (max - min) * 155;
  const positive = vals.at(-1) >= vals[0]; const color = positive ? '#67d9b1' : '#ff7e8b';
  svg.innerHTML = `<defs><linearGradient id="areaGradient" x1="0" x2="0" y1="0" y2="1"><stop offset="0%" stop-color="${color}" stop-opacity=".23"/><stop offset="100%" stop-color="${color}" stop-opacity="0"/></linearGradient></defs><polygon points="${firstX},${y0} ${coords.join(' ')} ${lastX},190 0,190" fill="url(#areaGradient)"/><polyline points="${coords.join(' ')}" fill="none" stroke="${color}" stroke-width="2" vector-effect="non-scaling-stroke" stroke-linecap="round" stroke-linejoin="round"/><circle cx="800" cy="${coords.at(-1).split(',')[1]}" r="3.5" fill="${color}"/>`;
}
function showResults(result) {
  const s = result.summary; $('#results').hidden = false; $('#live').hidden = true;
  $('#result-symbol').textContent = `${$('#symbol-select').value} · ${$('#timeframe').selectedOptions[0].textContent}`;
  const ret = $('#metric-return'); ret.textContent = `${s.returnPct >= 0 ? '+' : ''}${fmt(s.returnPct)}٪`; ret.className = s.returnPct >= 0 ? 'positive' : 'negative';
  $('#metric-final').textContent = `موجودی نهایی ${fmtMoney(s.finalBalance)}`;
  $('#metric-dd').textContent = `${fmt(s.maxDrawdownPct)}٪`; $('#metric-dd').className = 'negative';
  $('#metric-win').textContent = `${fmt(s.winRatePct)}٪`; $('#metric-win').className = s.winRatePct >= 50 ? 'positive' : '';
  $('#metric-trades').textContent = `${fmt(s.trades, 0)} معامله · ${fmt(s.wins, 0)} برد / ${fmt(s.losses, 0)} باخت`;
  $('#metric-pf').textContent = s.profitFactor === null ? '∞' : fmt(s.profitFactor); $('#metric-pf').className = s.profitFactor > 1 ? 'positive' : 'negative';
  drawChart(result.equity);
  $('#trades-body').innerHTML = result.trades.length ? result.trades.map(t => `<tr><td><span class="${t.side === 'خرید' ? 'buy' : 'sell'}">${t.side}</span></td><td>${fmt(t.entry, 4)}</td><td>${fmt(t.exit, 4)}</td><td class="${t.pnl >= 0 ? 'buy' : 'sell'}">${t.pnl >= 0 ? '+' : ''}${fmtMoney(t.pnl)}</td><td class="${t.pnlPct >= 0 ? 'buy' : 'sell'}">${t.pnlPct >= 0 ? '+' : ''}${fmt(t.pnlPct)}٪</td><td>${t.reason}</td></tr>`).join('') : '<tr><td colspan="6">در این بازه سیگنالی ثبت نشد.</td></tr>';
  $('#trade-caption').textContent = `${fmt(s.trades, 0)} معامله · ریسک ${fmt(s.riskPct)}٪ در هر معامله`;
  $('#result-disclaimer').textContent = result.disclaimer;
  $('#results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}
async function refreshIntelligence() {
  try {
    const params = new URLSearchParams({ symbol: $('#symbol-select').value, timeframe: $('#timeframe').value, provider: $('#provider').value, risk: $('#risk').value });
    const data = await api(`/api/intelligence?${params}`);
    $('#decision-direction').textContent = `${data.decision.direction} · ${data.trend}`;
    $('#decision-direction').className = `decision-direction ${data.decision.direction === 'خرید' ? 'positive' : data.decision.direction === 'فروش' ? 'negative' : ''}`;
    $('#decision-confidence').textContent = `قدرت اکتشافی ${fmt(data.decision.confidencePct, 0)}٪`;
    $('#confidence-fill').style.width = `${data.decision.confidencePct}%`;
    $('#layer-list').innerHTML = data.decision.layers.map(x => `<div><span>${escapeHtml(x.name)}</span><b>${escapeHtml(x.value)}</b></div>`).join('');
    $('#risk-output').textContent = `${fmt(data.risk.requestedRiskPct, 1)}٪ ← ${fmt(data.risk.adjustedRiskPct, 1)}٪`;
    $('#levels-output').textContent = `−${fmt(data.risk.stopLossPct)}٪ / +${fmt(data.risk.takeProfitPct)}٪`;
    $('#risk-note').textContent = `${data.risk.sizingBasis} · ATR فعلی ${fmt(data.atrPct)}٪`;
    $('#liquidity-score').textContent = `${fmt(data.liquidity.score, 0)}/۱۰۰`;
    $('#liquidity-label').textContent = data.liquidity.label;
    $('#liquidity-copy').textContent = `میانگین حجم ${fmt(data.liquidity.avgVolume, 1)} · مومنتوم ۵/۲۰ کندل ${fmt(data.momentum5Pct)}٪ / ${fmt(data.momentum20Pct)}٪`;
    $('#sentiment-label').textContent = `${data.sentiment.label} · ${data.sentiment.score}/۱۰۰`;
    const learned = data.learning;
    $('#learning-samples').textContent = `${fmt(learned.samples || 0, 0)} نمونه`;
    $('#learning-copy').textContent = learned.samples ? `بازده میانگین ۲۰ نمونه اخیر: ${fmt(learned.avgReturnPct)}٪ · افت میانگین: ${fmt(learned.avgDrawdownPct)}٪` : learned.status;
  } catch (e) { toast(e.message); }
}
async function scanMarkets() {
  const button = $('#scan-btn'); button.disabled = true; button.textContent = 'در حال اسکن…';
  try {
    const data = await api('/api/scan', { method: 'POST', body: JSON.stringify({ provider: $('#provider').value, timeframe: $('#timeframe').value, risk: $('#risk').value }) });
    $('#scanner-body').innerHTML = data.results.map(row => `<tr><td><b>${escapeHtml(row.symbol)}</b></td><td class="${row.direction === 'خرید' ? 'buy' : row.direction === 'فروش' ? 'sell' : ''}">${row.direction}</td><td>${row.trend}</td><td>${fmt(row.confidencePct, 0)}٪</td><td>${fmt(row.liquidityScore, 0)}/۱۰۰</td></tr>`).join('');
    toast(data.note);
  } catch (e) { toast(e.message); }
  finally { button.disabled = false; button.textContent = 'اسکن اکنون ←'; }
}
async function optimizeSelected() {
  const button = $('#optimize-btn'); button.disabled = true; button.textContent = 'در حال ارزیابی…';
  try {
    const data = await api('/api/optimize', { method: 'POST', body: JSON.stringify({ strategyId: $('#strategy-select').value, symbol: $('#symbol-select').value, timeframe: $('#timeframe').value, provider: $('#provider').value, risk: $('#risk').value, balance: $('#balance').value }) });
    const box = $('#optimizer-result'); box.textContent = `بهترین ترکیب از ${fmt(data.tested, 0)} حالت: ${JSON.stringify(data.best.params)} · بازده ${fmt(data.best.returnPct)}٪ · افت ${fmt(data.best.maxDrawdownPct)}٪ · ${fmt(data.best.trades, 0)} معامله. ${data.warning}`;
  } catch (e) { toast(e.message); }
  finally { button.disabled = false; button.textContent = 'ارزیابی ترکیب‌ها'; }
}
async function refreshLearning() {
  try {
    const summaries = await api('/api/learning');
    const total = summaries.reduce((sum, x) => sum + x.samples, 0);
    $('#learning-samples').textContent = `${fmt(total, 0)} نمونه`;
    if (total) $('#learning-copy').textContent = `${summaries.length} ستاپ دارای گزارش تاریخی است؛ امتیازها از نتایج ثبت‌شده‌اند و مدل یادگیری ماشین آموزش نمی‌بینند.`;
  } catch {}
}
async function run() {
  if (chosenMode() === 'paper') { $('#confirm-dialog').showModal(); return; }
  const btn = $('#run-btn'); btn.disabled = true; btn.textContent = 'در حال محاسبه…';
  try {
    const result = await api('/api/backtest', { method: 'POST', body: JSON.stringify({ strategyId: $('#strategy-select').value, symbol: $('#symbol-select').value, timeframe: $('#timeframe').value, risk: Number($('#risk').value), balance: Number($('#balance').value), provider: $('#provider').value }) });
    showResults(result); refreshLearning();
  } catch (e) { toast(e.message); }
  finally { btn.disabled = false; updateMode(); }
}
async function startPaper() {
  $('#confirm-dialog').close();
  try {
    const data = await api('/api/paper/start', { method: 'POST', body: JSON.stringify({ confirmed: true, strategyId: $('#strategy-select').value, symbol: $('#symbol-select').value, risk: Number($('#risk').value), balance: Number($('#balance').value) }) });
    $('#live-copy').textContent = `نشست ${data.id} · ${data.strategyName} روی ${data.symbol} با ریسک ${fmt(data.risk)}٪؛ این حالت فقط آزمایشی است و سفارش یا قیمت زنده دریافت نمی‌کند.`;
    $('#live').hidden = false; $('#results').hidden = true; $('#live').scrollIntoView({ behavior: 'smooth', block: 'center' }); toast('نشست آزمایشی شروع شد.');
  } catch (e) { toast(e.message); }
}
async function stopPaper() { try { await api('/api/paper/stop', { method: 'POST', body: '{}' }); $('#live').hidden = true; toast('نشست آزمایشی متوقف شد.'); } catch (e) { toast(e.message); } }
$('#risk').addEventListener('input', e => { $('#risk-value').textContent = `${fmt(e.target.value, 1)}٪`; });
document.querySelectorAll('input[name="mode"]').forEach(el => el.addEventListener('change', updateMode));
$('#strategy-select').addEventListener('change', updateDescription);
$('#run-btn').addEventListener('click', run); $('#cancel-paper').addEventListener('click', () => $('#confirm-dialog').close()); $('#confirm-paper').addEventListener('click', startPaper); $('#stop-paper').addEventListener('click', stopPaper);
$('#refresh-analysis').addEventListener('click', refreshIntelligence); $('#scan-btn').addEventListener('click', scanMarkets); $('#optimize-btn').addEventListener('click', optimizeSelected);
$('#symbol-select').addEventListener('change', refreshIntelligence); $('#timeframe').addEventListener('change', refreshIntelligence); $('#risk').addEventListener('change', refreshIntelligence);
$('#provider').addEventListener('change', refreshIntelligence);
$('#dismiss-results').addEventListener('click', () => { $('#results').hidden = true; });
$('#provider').addEventListener('change', e => { if (e.target.value === 'soft4fx') toast('اتصال مستقیم Soft4FX هنوز آماده نیست؛ در این نسخه از Binance یا داده آزمایشی استفاده کنید.'); if (e.target.value === 'binance' && ['XAUUSD','EURUSD'].includes($('#symbol-select').value)) $('#symbol-select').value = 'BTCUSDT'; });
$('#strategy-form').addEventListener('submit', async e => {
  e.preventDefault(); const message = $('#admin-message'); message.textContent = '';
  try { const item = await api('/api/strategies', { method: 'POST', body: JSON.stringify({ name: $('#new-name').value, type: $('#new-type').value, description: $('#new-description').value }) }); renderStrategies(await api('/api/strategies')); $('#strategy-select').value = item.id; updateDescription(); e.target.reset(); message.textContent = `«${item.name}» به کتابخانه اضافه شد.`; message.style.color = '#60d6ad'; toast('ستاپ جدید ثبت شد.'); }
  catch (err) { message.textContent = err.message; message.style.color = '#ff8290'; }
});
$('#today').textContent = new Intl.DateTimeFormat('fa-IR', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }).format(new Date());
api('/api/strategies').then(renderStrategies).catch(e => toast(e.message));
refreshIntelligence(); refreshLearning();
updateMode();
