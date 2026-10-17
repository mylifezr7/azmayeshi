class MockProvider {
  async getCandles({ symbol = 'BTCUSDT', limit = 700 }) {
    // Deterministic, deliberately synthetic candles. Never represent this as exchange history.
    let seed = [...symbol].reduce((n, c) => (n * 31 + c.charCodeAt(0)) >>> 0, 2026);
    const rand = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
    const candles = []; let close = symbol.includes('ETH') ? 3200 : symbol.includes('SOL') ? 150 : symbol.includes('XAU') ? 2350 : symbol.includes('EUR') ? 1.08 : 65000;
    for (let i = 0; i < limit; i++) {
      const regime = Math.sin(i / 47) * 0.00048;
      const change = regime + (rand() - 0.51) * 0.014;
      const open = close; close = Math.max(close * 0.3, close * (1 + change));
      const spread = close * (0.001 + rand() * 0.006);
      candles.push({ time: Date.now() - (limit - i) * 3600000, open, high: Math.max(open, close) + spread * rand(), low: Math.min(open, close) - spread * rand(), close, volume: 10 + rand() * 100 });
    }
    return candles;
  }
}
class BinanceProvider {
  async getCandles({ symbol = 'BTCUSDT', timeframe = '1h', limit = 700 }) {
    const allowed = new Set(['1m','3m','5m','15m','30m','1h','2h','4h','6h','8h','12h','1d','3d','1w']);
    if (!/^[A-Z0-9]{5,20}$/.test(symbol) || !allowed.has(timeframe)) throw new Error('نماد یا بازه زمانی برای Binance معتبر نیست.');
    const response = await fetch(`https://api.binance.com/api/v3/klines?symbol=${encodeURIComponent(symbol)}&interval=${timeframe}&limit=${Math.min(1000, Math.max(100, Number(limit) || 700))}`);
    if (!response.ok) throw new Error(`دریافت اطلاعات Binance ناموفق بود (${response.status}).`);
    const rows = await response.json();
    return rows.map(k => ({ time: k[0], open: +k[1], high: +k[2], low: +k[3], close: +k[4], volume: +k[5] }));
  }
}
class Soft4FXProvider {
  async getCandles() { throw new Error('اتصال مستقیم Soft4FX هنوز پیاده‌سازی نشده است؛ فایل CSV در نقشه راه قرار دارد.'); }
}
function createProvider(id) {
  if (id === 'mock') return new MockProvider();
  if (id === 'binance') return new BinanceProvider();
  if (id === 'soft4fx') return new Soft4FXProvider();
  throw new Error('منبع داده ناشناخته است.');
}
module.exports = { createProvider, MockProvider, BinanceProvider, Soft4FXProvider };
