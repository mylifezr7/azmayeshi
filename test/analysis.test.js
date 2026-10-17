const test=require('node:test');const assert=require('node:assert/strict');
const {analyze,rsi}=require('../src/analysis');const {marketClock}=require('../src/market-clock');const {DemoProvider,rankUniverse}=require('../src/providers');const {runBacktest}=require('../src/backtest');
const demo=new DemoProvider();
test('crypto demo analysis returns technical, flow, risk levels, and clearly empty real fundamentals/news',async()=>{
 const symbol='BTCUSDT',candles=await demo.getCandles({symbol,timeframe:'1h',limit:260});
 const [ticker,orderBook,aggTrades,fundamentals,news]=await Promise.all([demo.getTicker(symbol),demo.getOrderBook(symbol),demo.getAggTrades(symbol),demo.getFundamentals([symbol]),demo.getNews(symbol)]);
 const result=analyze({symbol,candles,ticker,orderBook,aggTrades,fundamentals:fundamentals[symbol],news,riskPct:1,balance:10000});
 assert.equal(result.symbol,symbol);assert.ok(result.technical.atr>0);assert.ok(result.technical.rsi>=0&&result.technical.rsi<=100);
 assert.ok(result.plan.stopPct>=.5&&result.plan.stopPct<=8);assert.equal(result.plan.targetPct,result.plan.stopPct*2);
 assert.ok(result.pressure.takerImbalancePct>=-100&&result.pressure.takerImbalancePct<=100);
 assert.equal(result.fundamentals.available,false);assert.equal(result.news.available,false);
});
test('demo candles respect the selected interval',async()=>{const c=await demo.getCandles({symbol:'ETHUSDT',timeframe:'4h',limit:60});assert.equal(c[1].time-c[0].time,4*60*60*1000);});
test('crypto exchange is always open and scheduler clock follows local timezone',()=>{const clock=marketClock(new Date('2026-10-19T14:00:00.000Z'));assert.equal(clock.marketOpen,true);assert.match(clock.market,/شبانه‌روزی/);assert.ok(clock.timeZone);});
test('volume universe contains BTC plus ten other demo markets',async()=>{const assets=await demo.getUniverse();assert.equal(assets.length,11);assert.equal(assets[0].symbol,'BTCUSDT');assert.equal(new Set(assets.map(x=>x.symbol)).size,11);});
test('universe ranks BTC plus ten highest-volume eligible USDT pairs',()=>{const rows=[{symbol:'BTCUSDT',quoteVolume:'10',lastPrice:'1',priceChangePercent:'0'},{symbol:'USDCUSDT',quoteVolume:'9999'},{symbol:'BTCUPUSDT',quoteVolume:'9998'},...Array.from({length:12},(_,i)=>({symbol:`COIN${String.fromCharCode(65+i)}USDT`,quoteVolume:String(100-i),lastPrice:'2',priceChangePercent:'1'}))];const ranked=rankUniverse(rows);assert.equal(ranked.length,11);assert.equal(ranked[0].symbol,'BTCUSDT');assert.equal(ranked[1].symbol,'COIN AUSDT'.replace(' ',''));assert.ok(!ranked.some(x=>x.symbol==='USDCUSDT'||x.symbol==='BTCUPUSDT'));});
test('multi-factor historical backtest returns risk and performance metrics',async()=>{const candles=await demo.getCandles({symbol:'BTCUSDT',timeframe:'1h',limit:500});const result=runBacktest({symbol:'BTCUSDT',candles,riskPct:1,initialBalance:10000});assert.ok(Number.isFinite(result.summary.returnPct));assert.ok(result.summary.finalBalance>0);assert.equal(result.summary.riskPct,1);assert.equal(result.candlesCount,500);});
test('RSI remains bounded on falling prices',()=>{assert.equal(rsi(Array.from({length:40},(_,i)=>100-i),14),0);});
