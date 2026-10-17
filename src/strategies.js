const fs = require('node:fs');
const path = require('node:path');
const STORE = path.join(__dirname, 'custom-strategies.json');
const builtins = [
  { id: 'ema-cross', name: 'تقاطع میانگین متحرک', short: 'EMA Cross', category: 'روندی', description: 'ورود با تقاطع EMA سریع و کند؛ خروج با تقاطع معکوس.', params: { fast: 12, slow: 26 }, type: 'ema-cross', color: '#7c6cf2', riskHint: 'بازارهای رونددار' },
  { id: 'rsi-reversal', name: 'بازگشت از اشباع RSI', short: 'RSI Reversal', category: 'بازگشتی', description: 'خرید پس از خروج RSI از اشباع فروش و فروش در اشباع خرید.', params: { period: 14, oversold: 30, overbought: 70 }, type: 'rsi-reversal', color: '#28c4a0', riskHint: 'بازارهای رنج' },
  { id: 'breakout', name: 'شکست کانال', short: 'Channel Breakout', category: 'مومنتوم', description: 'ورود با شکست سقف یا کف محدوده کندل‌های اخیر.', params: { lookback: 20 }, type: 'breakout', color: '#f4a340', riskHint: 'حرکت‌های پرقدرت' }
];
function custom() { try { return JSON.parse(fs.readFileSync(STORE, 'utf8')); } catch { return []; } }
function list() { return [...builtins, ...custom()]; }
function getStrategy(id) { const result = list().find(x => x.id === id); if (!result) { const e = new Error('استراتژی انتخاب‌شده پیدا نشد'); e.status = 404; throw e; } return result; }
function addStrategy(input) {
  const name = String(input.name || '').trim();
  const type = String(input.type || 'ema-cross');
  if (name.length < 3 || name.length > 50) throw new Error('نام ستاپ باید بین ۳ تا ۵۰ کاراکتر باشد.');
  if (!['ema-cross', 'rsi-reversal', 'breakout'].includes(type)) throw new Error('نوع ستاپ پشتیبانی نمی‌شود.');
  const items = custom();
  const id = `custom-${Date.now().toString(36)}`;
  const item = { id, name, short: name, category: 'سفارشی', description: String(input.description || 'ستاپ سفارشی مبتنی بر الگوی انتخاب‌شده.'), params: type === 'ema-cross' ? { fast: 10, slow: 30 } : type === 'rsi-reversal' ? { period: 14, oversold: 30, overbought: 70 } : { lookback: 20 }, type, color: '#9b8cff', riskHint: 'نیازمند ارزیابی مستقل' };
  items.push(item); fs.writeFileSync(STORE, JSON.stringify(items, null, 2)); return item;
}
module.exports = { strategies: { list }, getStrategy, addStrategy };
