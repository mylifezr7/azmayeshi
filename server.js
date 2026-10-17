const http=require('node:http');const fs=require('node:fs');const path=require('node:path');const {URL}=require('node:url');
const {createProvider}=require('./src/providers');const {analyze}=require('./src/analysis');const {runBacktest}=require('./src/backtest');const {marketClock}=require('./src/market-clock');const paper=require('./src/paper');
const ROOT=__dirname,PORT=Number(process.env.PORT||3000);
function json(res,status,body){res.writeHead(status,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify(body));}
function serveStatic(res,pathname){const requested=pathname==='/'?'/public/index.html':`/public${pathname}`,file=path.resolve(ROOT,`.${requested}`);if(!file.startsWith(path.join(ROOT,'public')+path.sep))return json(res,403,{error:'دسترسی غیرمجاز'});fs.readFile(file,(err,content)=>{if(err)return json(res,404,{error:'فایل پیدا نشد'});const ext=path.extname(file),type=ext==='.css'?'text/css':ext==='.js'?'text/javascript':'text/html';res.writeHead(200,{'Content-Type':`${type}; charset=utf-8`});res.end(content);});}
async function mapLimit(items,limit,fn){const out=Array(items.length);let cursor=0;await Promise.all(Array.from({length:Math.min(limit,items.length)},async()=>{while(true){const i=cursor++;if(i>=items.length)return;try{out[i]=await fn(items[i],i);}catch(error){out[i]={symbol:items[i].symbol,error:error.message,quoteVolume24h:items[i].quoteVolume};}}}));return out;}
async function analyzeAsset(provider,symbol,{timeframe='1h',risk=1,balance=10000,limit=500}={}){
 const candles=await provider.getCandles({symbol,timeframe,limit});
 const [ticker,orderBook,aggTrades,fundamentalsBySymbol,news]=await Promise.allSettled([
  provider.getTicker(symbol),provider.getOrderBook(symbol),provider.getAggTrades(symbol),provider.getFundamentals([symbol]),provider.getNews(symbol)
 ]);
 const fundamentals=fundamentalsBySymbol.status==='fulfilled'?fundamentalsBySymbol.value[symbol]:null;
 const newsData=news.status==='fulfilled'?news.value:{items:[],status:news.reason?.message||'منبع خبر دردسترس نیست.'};
 const result=analyze({symbol,candles,ticker:ticker.status==='fulfilled'?ticker.value:{price:candles.at(-1).close,source:'آخرین کندل'},orderBook:orderBook.status==='fulfilled'?orderBook.value:null,aggTrades:aggTrades.status==='fulfilled'?aggTrades.value:null,fundamentals,news:newsData,riskPct:risk,balance,timeframe});
 return{...result,candles};
}
const server=http.createServer(async(req,res)=>{
 const url=new URL(req.url,`http://${req.headers.host||'localhost'}`);
 try{
  if(req.method==='GET'&&url.pathname==='/api/health')return json(res,200,{ok:true,market:'crypto-24/7',provider:'Binance public market data',realOrders:false,news:Boolean(process.env.CRYPTOPANIC_API_KEY)});
  if(req.method==='GET'&&url.pathname==='/api/market-clock')return json(res,200,marketClock());
  if(req.method==='GET'&&url.pathname==='/api/paper/status'){return json(res,200,await paper.monitor(createProvider('binance')));}
  if(req.method==='POST'&&url.pathname==='/api/paper/close-all'){for await(const _chunk of req){}return json(res,200,paper.closeAll());}
  if(req.method==='GET'&&url.pathname==='/api/universe'){
   const provider=createProvider(url.searchParams.get('provider')||'binance');const assets=await provider.getUniverse();return json(res,200,{assets,source:(url.searchParams.get('provider')||'binance')==='demo'?'دادهٔ نمایشی':'حجم معاملات ۲۴ساعتهٔ جفت‌های USDT در Binance',updatedAt:new Date().toISOString(),note:'BTC به‌علاوهٔ ۱۰ جفت دیگر با بیشترین حجم Quote Volume؛ استیبل‌کوین‌ها و توکن‌های اهرمی حذف شده‌اند.'});
  }
  if(req.method==='GET'&&url.pathname==='/api/analysis'){
   const symbol=(url.searchParams.get('symbol')||'BTCUSDT').toUpperCase();if(!/^[A-Z0-9]{2,16}USDT$/.test(symbol))return json(res,400,{error:'نماد باید جفت معتبر USDT باشد؛ نمونه BTCUSDT.'});
   const provider=createProvider(url.searchParams.get('provider')||'binance');const result=await analyzeAsset(provider,symbol,{timeframe:url.searchParams.get('timeframe')||'1h',risk:url.searchParams.get('risk'),balance:url.searchParams.get('balance')});delete result.candles;return json(res,200,result);
  }
  if(req.method==='POST'&&url.pathname==='/api/portfolio-run'){
   let body='';for await(const chunk of req){body+=chunk;if(body.length>100_000)throw new Error('درخواست خیلی بزرگ است.');}const input=body?JSON.parse(body):{};
   const providerId=input.provider||'binance',provider=createProvider(providerId),timeframe=input.timeframe||'1h',risk=Math.min(3,Math.max(.1,Number(input.risk)||1)),balance=Math.max(100,Number(input.balance)||10000);if(input.paper&&providerId!=='binance')throw new Error('Paper trade فقط با قیمت Binance و دادهٔ بازار قابل اجراست؛ حالت دمو را خاموش کن.');
   const universe=await provider.getUniverse(),symbols=universe.map(x=>x.symbol),fundamentals=await provider.getFundamentals(symbols);
   const rows=await mapLimit(universe,3,async asset=>{
    const candles=await provider.getCandles({symbol:asset.symbol,timeframe,limit:500});
    const [ticker,book,flow,news]=await Promise.allSettled([provider.getTicker(asset.symbol),provider.getOrderBook(asset.symbol),provider.getAggTrades(asset.symbol),provider.getNews(asset.symbol)]);
    const context=analyze({symbol:asset.symbol,candles,ticker:ticker.status==='fulfilled'?ticker.value:{price:candles.at(-1).close},orderBook:book.status==='fulfilled'?book.value:null,aggTrades:flow.status==='fulfilled'?flow.value:null,fundamentals:fundamentals[asset.symbol]||null,news:news.status==='fulfilled'?news.value:{items:[]},riskPct:risk,balance,timeframe});
    const backtest=runBacktest({symbol:asset.symbol,candles,riskPct:risk,initialBalance:balance,feeRate:input.feeRate});
    return{symbol:asset.symbol,volumeRank:asset.volumeRank,quoteVolume24h:asset.quoteVolume,change24hPct:asset.change24hPct,analysis:context,backtest:backtest.summary,candlesCount:backtest.candlesCount,from:backtest.from,to:backtest.to};
   });
   const paperResult=input.paper?paper.openPaperTrades(rows,{balance,riskPct:risk}):null;
   return json(res,200,{source:providerId==='binance'?'Binance · بازار واقعی':'دادهٔ مصنوعی دمو',timeframe,riskPct:risk,balance,universe:universe.map(x=>({symbol:x.symbol,quoteVolume24h:x.quoteVolume,change24hPct:x.change24hPct,volumeRank:x.volumeRank})),results:rows,paper:paperResult,completedAt:new Date().toISOString(),disclaimer:'بک‌تست تاریخی شامل EMA/MACD/RSI، شکست کانال، بازگشت Fib، فشار taker داخل کندل و حدهای ATR است. تاریخچهٔ دفتر سفارش، خبر و فاندامنتال برای این بازه در API عمومی موجود نیست؛ آن‌ها فقط در فیلتر Paper زنده استفاده می‌شوند. کارمزد فرضی ۰.۱٪ هر سمت است؛ لغزش/فاندینگ/تأخیر مدل نشده‌اند. نتایج آینده را تضمین نمی‌کنند.'});
  }
  if(req.method==='GET')return serveStatic(res,url.pathname);
  return json(res,404,{error:'مسیر پیدا نشد'});
 }catch(error){return json(res,error.status||400,{error:error.message||'خطای غیرمنتظره'});}
});
server.listen(PORT,'0.0.0.0',()=>console.log(`Crypto analysis listening on 0.0.0.0:${PORT}`));
