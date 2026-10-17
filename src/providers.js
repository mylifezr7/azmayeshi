const BINANCE_BASE = (process.env.BINANCE_API_BASE || 'https://data-api.binance.vision/api/v3').replace(/\/$/, '');
const DAY = 86400;
const STABLES = new Set(['USDC','FDUSD','TUSD','USDP','DAI','USDE','BUSD','USTC','EUR','AEUR','USDD','PYUSD']);
const COIN_IDS = { BTC:'bitcoin', ETH:'ethereum', BNB:'binancecoin', SOL:'solana', XRP:'ripple', DOGE:'dogecoin', TRX:'tron', ADA:'cardano', LINK:'chainlink', LTC:'litecoin', AVAX:'avalanche-2', BCH:'bitcoin-cash', TON:'the-open-network', XLM:'stellar', SUI:'sui', HBAR:'hedera-hashgraph', SHIB:'shiba-inu', DOT:'polkadot', UNI:'uniswap', AAVE:'aave', ETC:'ethereum-classic', ICP:'internet-computer', NEAR:'near', FIL:'filecoin', ATOM:'cosmos', ARB:'arbitrum', OP:'optimism', APT:'aptos', WLD:'worldcoin', HYPE:'hyperliquid', PEPE:'pepe', WIF:'dogwifcoin', FET:'artificial-superintelligence-alliance', RENDER:'render-token', INJ:'injective-protocol', SEI:'sei-network', TAO:'bittensor', ONDO:'ondo-finance', POL:'polygon-ecosystem-token', VET:'vechain', ALGO:'algorand', RUNE:'thorchain', CRV:'curve-dao-token', PAXG:'pax-gold' };
let cgCache = { at: 0, rows: null };

async function fetchJson(url, label, timeout = 12000) {
  const response = await fetch(url, { signal: AbortSignal.timeout(timeout), headers: { 'Accept': 'application/json', 'User-Agent': 'MarketNegar/1.0' } });
  if (!response.ok) throw new Error(`${label}: پاسخ HTTP ${response.status}`);
  return response.json();
}
function rankUniverse(rows) {
  const candidates = rows.filter(x => x.symbol?.endsWith('USDT') && Number(x.quoteVolume) > 0 && x.symbol.length <= 18).filter(x => {
    const base = x.symbol.slice(0, -4);
    return !STABLES.has(base) && !/(UP|DOWN|BULL|BEAR|[235]L|[235]S)$/.test(base);
  }).sort((a,b) => Number(b.quoteVolume) - Number(a.quoteVolume));
  const bitcoin=candidates.find(x=>x.symbol==='BTCUSDT'),others=candidates.filter(x=>x.symbol!=='BTCUSDT').slice(0,10);
  if(!bitcoin)throw new Error('نماد BTCUSDT از Binance دریافت نشد.');
  return[bitcoin,...others].map((x,i)=>({symbol:x.symbol,base:x.symbol.slice(0,-4),price:Number(x.lastPrice),quoteVolume:Number(x.quoteVolume),change24hPct:Number(x.priceChangePercent),volumeRank:x.symbol==='BTCUSDT'?'BTC':i}));
}
class BinanceProvider {
  async request(route) { return fetchJson(`${BINANCE_BASE}${route}`, 'Binance'); }
  async getUniverse() { return rankUniverse(await this.request('/ticker/24hr')); }
  async getTicker(symbol) {
    const q = await this.request(`/ticker/24hr?symbol=${encodeURIComponent(symbol.toUpperCase())}`);
    return { price: Number(q.lastPrice), previousClose: Number(q.prevClosePrice), changePct: Number(q.priceChangePercent), quoteVolume: Number(q.quoteVolume), source: 'Binance · ۲۴ساعت', updatedAt: new Date().toISOString() };
  }
  async getCandles({ symbol, timeframe = '1h', limit = 500 }) {
    const intervals = new Set(['1m','3m','5m','15m','30m','1h','2h','4h','6h','8h','12h','1d','3d','1w']);
    if (!intervals.has(timeframe)) throw new Error('بازهٔ زمانی Binance معتبر نیست.');
    const safeLimit = Math.max(100, Math.min(1000, Number(limit) || 500));
    const q = new URLSearchParams({ symbol: symbol.toUpperCase(), interval: timeframe, limit: String(safeLimit) });
    const rows = await this.request(`/klines?${q}`);
    return rows.map(k => ({ time: k[0], open: +k[1], high: +k[2], low: +k[3], close: +k[4], volume: +k[5], quoteVolume: +k[7], trades: +k[8], takerBuyBaseVolume: +k[9], takerBuyQuoteVolume: +k[10] }));
  }
  async getOrderBook(symbol) {
    const q = new URLSearchParams({ symbol: symbol.toUpperCase(), limit: '100' });
    const book = await this.request(`/depth?${q}`);
    const bidUsd = book.bids.reduce((s,[p,v]) => s + Number(p)*Number(v), 0);
    const askUsd = book.asks.reduce((s,[p,v]) => s + Number(p)*Number(v), 0);
    const total = bidUsd + askUsd;
    return { bidNotionalUsd: bidUsd, askNotionalUsd: askUsd, imbalancePct: total ? (bidUsd-askUsd)/total*100 : 0, bids: book.bids.length, asks: book.asks.length, source: 'Binance order book · ۱۰۰ سطح' };
  }
  async getAggTrades(symbol) {
    const q = new URLSearchParams({ symbol: symbol.toUpperCase(), limit: '500' });
    const rows = await this.request(`/aggTrades?${q}`);
    let buyUsd=0, sellUsd=0, largeBuy=0, largeSell=0, whaleBuy=0, whaleSell=0;
    for (const t of rows) {
      const notional = Number(t.p)*Number(t.q);
      if (t.m) sellUsd += notional; else buyUsd += notional;
      if (notional >= 10000) { if(t.m) largeSell++; else largeBuy++; }
      if (notional >= 100000) { if(t.m) whaleSell++; else whaleBuy++; }
    }
    const total=buyUsd+sellUsd;
    return { buyNotionalUsd: buyUsd, sellNotionalUsd: sellUsd, takerImbalancePct: total ? (buyUsd-sellUsd)/total*100 : 0, largeBuyTrades: largeBuy, largeSellTrades: largeSell, veryLargeBuyTrades: whaleBuy, veryLargeSellTrades: whaleSell, sampleTrades: rows.length, source: 'Binance aggregate trades · آخرین ۵۰۰' };
  }
  async getFundamentals(symbols) {
    try {
      if (!cgCache.rows || Date.now()-cgCache.at > 10*60_000) {
        const query = new URLSearchParams({ vs_currency: 'usd', order: 'market_cap_desc', per_page: '250', page: '1', sparkline: 'false', price_change_percentage: '24h,7d,30d' });
        cgCache.rows = await fetchJson(`https://api.coingecko.com/api/v3/coins/markets?${query}`, 'CoinGecko', 15000); cgCache.at=Date.now();
      }
      const rows = {};
      for (const symbol of symbols) {
        const base = symbol.replace(/USDT$/, '').replace(/^1000/, '');
        const id = COIN_IDS[base];
        const coin = id ? cgCache.rows.find(x => x.id === id) : cgCache.rows.find(x => x.symbol?.toUpperCase() === base);
        if (!coin) { rows[symbol] = null; continue; }
        rows[symbol] = { name: coin.name, marketCapUsd: coin.market_cap, marketCapRank: coin.market_cap_rank, fullyDilutedValuationUsd: coin.fully_diluted_valuation, totalVolume24hUsd: coin.total_volume, circulatingSupply: coin.circulating_supply, totalSupply: coin.total_supply, maxSupply: coin.max_supply, supplyFloatPct: coin.total_supply ? coin.circulating_supply/coin.total_supply*100 : null, priceChange24hPct: coin.price_change_percentage_24h, priceChange7dPct: coin.price_change_percentage_7d_in_currency, priceChange30dPct: coin.price_change_percentage_30d_in_currency, source: 'CoinGecko' };
      }
      return rows;
    } catch { return Object.fromEntries(symbols.map(s => [s, null])); }
  }
  async getNews(symbol) {
    const token = process.env.CRYPTOPANIC_API_KEY;
    if (!token) return { items: [], status: 'برای خبر واقعی CRYPTOPANIC_API_KEY تنظیم نشده است.' };
    try {
      const base = symbol.replace(/USDT$/, '');
      const q = new URLSearchParams({ auth_token: token, currencies: base, public: 'true' });
      const result = await fetchJson(`https://cryptopanic.com/api/v1/posts/?${q}`, 'CryptoPanic', 15000);
      const items=(result.results||[]).slice(0,8).map(x => {
        const votes=x.votes||{}; const diff=Number(votes.positive||0)-Number(votes.negative||0);
        const sentiment=diff>0?'مثبت (رأی کاربران)':diff<0?'منفی (رأی کاربران)':'خنثی / بدون رأی';
        return { headline:x.title, url:x.url, source:x.source?.title||'CryptoPanic', time:x.published_at, sentiment, votes:diff };
      });
      return { items, status: items.length?'خبرهای تازهٔ بازیابی‌شده':'خبر تازه‌ای برای این نماد نیست.' };
    } catch(error) { return { items: [], status: error.message }; }
  }
}
class DemoProvider {
  async getCandles({ symbol='BTCUSDT', timeframe='1h', limit=500 }) {
    let seed=[...symbol].reduce((n,c)=>(n*31+c.charCodeAt(0))>>>0,2026); const rand=()=>{seed=(seed*1664525+1013904223)>>>0;return seed/4294967296;};
    let close=symbol.startsWith('BTC')?68000:symbol.startsWith('ETH')?3500:150+rand()*300; const interval=({ '15m':900000,'1h':3600000,'4h':14400000,'1d':86400000 })[timeframe]||3600000; const candles=[];
    for(let i=0;i<Math.min(limit,800);i++){const open=close;close*=1+Math.sin(i/42)*.0007+(rand()-.51)*.018;const range=close*(.002+rand()*.012);const volume=1+rand()*20;const buy=rand();candles.push({time:Date.now()-(limit-i)*interval,open,high:Math.max(open,close)+range*rand(),low:Math.min(open,close)-range*rand(),close,volume,takerBuyBaseVolume:volume*buy,takerBuyQuoteVolume:volume*buy*close,quoteVolume:volume*close});}
    return candles;
  }
  async getUniverse(){return ['BTCUSDT','ETHUSDT','SOLUSDT','BNBUSDT','XRPUSDT','DOGEUSDT','ADAUSDT','TRXUSDT','LINKUSDT','AVAXUSDT','LTCUSDT'].map((symbol,i)=>({symbol,base:symbol.slice(0,-4),price:100+i*100,quoteVolume:(11-i)*1e9,change24hPct:(i%2?1:-1)*i,volumeRank:i||'BTC'}));}
  async getTicker(symbol){return {price: symbol==='BTCUSDT'?68000:250,previousClose:248,changePct:.8,source:'دادهٔ مصنوعی دمو',quoteVolume:1e9};}
  async getOrderBook(){return {bidNotionalUsd:5000000,askNotionalUsd:4500000,imbalancePct:5.26,bids:100,asks:100,source:'نمونهٔ دمو؛ نه order book واقعی'};}
  async getAggTrades(){return {buyNotionalUsd:800000,sellNotionalUsd:700000,takerImbalancePct:6.67,largeBuyTrades:5,largeSellTrades:3,veryLargeBuyTrades:1,veryLargeSellTrades:0,sampleTrades:500,source:'نمونهٔ دمو؛ نه معاملات واقعی'};}
  async getFundamentals(symbols){return Object.fromEntries(symbols.map(s=>[s,null]));}
  async getNews(){return {items:[],status:'در حالت دمو خبر واقعی نمایش داده نمی‌شود.'};}
}
function createProvider(id='binance'){if(id==='binance')return new BinanceProvider();if(id==='demo')return new DemoProvider();throw new Error('منبع داده ناشناخته است.');}
module.exports={createProvider,BinanceProvider,DemoProvider,rankUniverse};
