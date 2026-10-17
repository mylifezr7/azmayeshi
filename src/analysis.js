const { marketClock } = require('./market-clock');
function ema(values, period) {
  const out=Array(values.length).fill(null), alpha=2/(period+1); let state=null;
  values.forEach((v,i)=>{state=state===null?v:v*alpha+state*(1-alpha);if(i>=period-1)out[i]=state;});
  return out;
}
function rsi(values, period=14) {
  if(values.length<=period)return null;
  let up=0,down=0;
  for(let i=1;i<=period;i++){const d=values[i]-values[i-1];up+=Math.max(0,d);down+=Math.max(0,-d);}
  let avgUp=up/period,avgDown=down/period;
  for(let i=period+1;i<values.length;i++){const d=values[i]-values[i-1];avgUp=(avgUp*(period-1)+Math.max(0,d))/period;avgDown=(avgDown*(period-1)+Math.max(0,-d))/period;}
  return avgDown===0?100:100-100/(1+avgUp/avgDown);
}
function marketStructure(candles, trendUp, atr) {
  const current=candles.at(-1),previous=candles.at(-2), prior=candles.slice(-21,-1);
  const channelHigh=Math.max(...prior.map(c=>c.high)),channelLow=Math.min(...prior.map(c=>c.low));
  const swing=candles.slice(-101,-1),swingHigh=Math.max(...swing.map(c=>c.high)),swingLow=Math.min(...swing.map(c=>c.low)),range=swingHigh-swingLow;
  const ratios=[.382,.5,.618];
  const levels=trendUp?ratios.map(r=>({ratio:r,price:swingHigh-range*r})):ratios.map(r=>({ratio:r,price:swingLow+range*r}));
  const tolerance=Math.max(atr*.22,current.close*.0015);
  const fibLongBounce=trendUp&&levels.some(x=>current.low<=x.price+tolerance&&current.close>x.price&&previous.close<=x.price);
  const fibShortReject=!trendUp&&levels.some(x=>current.high>=x.price-tolerance&&current.close<x.price&&previous.close>=x.price);
  const channelLongBreak=current.close>channelHigh,channelShortBreak=current.close<channelLow;
  return{channelHigh,channelLow,channelLongBreak,channelShortBreak,swingHigh,swingLow,fibonacciLevels:levels,fibLongBounce,fibShortReject,tolerance};
}
function analyze({symbol,candles,ticker={},orderBook=null,aggTrades=null,fundamentals=null,news={items:[]},riskPct=1,balance=10000,timeframe='1h'}) {
  if(!Array.isArray(candles)||candles.length<100)throw new Error('برای تحلیل و فیبوناچی حداقل ۱۰۰ کندل لازم است.');
  const closes=candles.map(c=>c.close),last=candles.at(-1),price=Number(ticker.price||last.close),emaFast=ema(closes,12),emaSlow=ema(closes,26),macd=closes.map((_,i)=>emaFast[i]!=null&&emaSlow[i]!=null?emaFast[i]-emaSlow[i]:0),macdSignal=ema(macd,9),macdHist=macd.at(-1)-macdSignal.at(-1),rsiValue=rsi(closes,14),trendUp=emaFast.at(-1)>emaSlow.at(-1);
  const ranges=[];for(let i=Math.max(1,candles.length-14);i<candles.length;i++){const prev=closes[i-1];ranges.push(Math.max(candles[i].high-candles[i].low,Math.abs(candles[i].high-prev),Math.abs(candles[i].low-prev)));}
  const atr=ranges.reduce((s,x)=>s+x,0)/ranges.length,structure=marketStructure(candles,trendUp,atr);
  let historicalBuy=0,historicalSell=0;for(const c of candles.slice(-20)){const buy=Number.isFinite(c.takerBuyQuoteVolume)?c.takerBuyQuoteVolume:(c.quoteVolume||c.volume||0)*Math.max(0,Math.min(1,(c.close-c.low)/Math.max(1e-12,c.high-c.low)));historicalBuy+=buy;historicalSell+=Number.isFinite(c.quoteVolume)?Math.max(0,c.quoteVolume-buy):(c.quoteVolume||c.volume||0)-buy;}
  const historicalTotal=historicalBuy+historicalSell,historicalPressure=historicalTotal?(historicalBuy-historicalSell)/historicalTotal*100:0,tradePressure=aggTrades?.takerImbalancePct??historicalPressure,bookPressure=orderBook?.imbalancePct??0;
  const technical=[{label:'روند EMA ۱۲/۲۶',value:trendUp?'صعودی':'نزولی',score:trendUp?1:-1},{label:'MACD',value:macdHist>=0?'مثبت':'منفی',score:macdHist>=0?1:-1},{label:'RSI ۱۴',value:rsiValue.toFixed(1),score:rsiValue<35?1:rsiValue>70?-1:0},{label:'شکست کانال ۲۰ کندل',value:structure.channelLongBreak?'شکست رو به بالا':structure.channelShortBreak?'شکست رو به پایین':'بدون شکست',score:structure.channelLongBreak?1:structure.channelShortBreak?-1:0},{label:'بازگشت فیبوناچی',value:structure.fibLongBounce?'واکنش صعودی':structure.fibShortReject?'رد نزولی':'بدون واکنش',score:structure.fibLongBounce?1:structure.fibShortReject?-1:0}];
  const technicalScore=technical.reduce((s,x)=>s+x.score,0),flowScore=tradePressure>10?1:tradePressure< -10?-1:0,bookScore=bookPressure>12?1:bookPressure< -12?-1:0;
  let fundamentalScore=0;const fundamentalFactors=[];
  if(fundamentals?.priceChange7dPct!=null){const v=Number(fundamentals.priceChange7dPct);fundamentalScore+=v>5?1:v< -5?-1:0;fundamentalFactors.push({label:'تغییر بازار ۷روزه',value:`${v.toFixed(2)}٪`});}
  if(fundamentals?.supplyFloatPct!=null){const v=Number(fundamentals.supplyFloatPct);if(v<35)fundamentalScore--;fundamentalFactors.push({label:'عرضه در گردش',value:`${v.toFixed(1)}٪`,risk:v<35?'رقیق‌شدن عرضه':'—'});}
  if(fundamentals?.marketCapUsd&&fundamentals?.fullyDilutedValuationUsd){const ratio=fundamentals.fullyDilutedValuationUsd/fundamentals.marketCapUsd;if(ratio>2){fundamentalScore--;fundamentalFactors.push({label:'FDV / ارزش بازار',value:ratio.toFixed(2),risk:'فاصله عرضه بالا'});}else fundamentalFactors.push({label:'FDV / ارزش بازار',value:ratio.toFixed(2)});}
  const newsItems=(news?.items||[]).filter(item=>{const time=Date.parse(item.time||'');return Number.isFinite(time)&&time<=Date.now()&&Date.now()-time<24*60*60*1000;}),newsScore=newsItems.reduce((s,item)=>s+(String(item.sentiment).startsWith('مثبت')?1:String(item.sentiment).startsWith('منفی')?-1:0),0),newsAvailable=newsItems.length>0;
  const newsLayer=Math.max(-1,Math.min(1,newsScore)),combined=technicalScore+flowScore+bookScore+fundamentalScore+newsLayer;
  const positiveStructure=structure.channelLongBreak||structure.fibLongBounce,negativeStructure=structure.channelShortBreak||structure.fibShortReject;
  const fundamentalReady=Boolean(fundamentals)&&(fundamentalScore>=0),longConditions={trend:trendUp,structure:positiveStructure,takerBuy:tradePressure>0,orderBookBuy:bookPressure>0,fundamental:fundamentalReady,news:newsAvailable&&newsScore>0};
  const shortConditions={trend:!trendUp,structure:negativeStructure,takerSell:tradePressure<0,orderBookSell:bookPressure<0,fundamental:Boolean(fundamentals)&&fundamentalScore<=0,news:newsAvailable&&newsScore<0};
  const canLong=Object.values(longConditions).every(Boolean),canShort=Object.values(shortConditions).every(Boolean),candidateSide=canLong?'long':canShort?'short':combined<0?'short':'long';
  const direction=canLong?'ورود لانگ · همه تأییدها':canShort?'ورود شورت · همه تأییدها':combined>=3?'تمایل صعودی · ورود تأیید نشد':combined<=-3?'تمایل نزولی · ورود تأیید نشد':'خنثی / انتظار';
  const stopPct=Math.min(8,Math.max(.5,atr/price*1.5*100)),distance=price*stopPct/100,risk=Math.min(3,Math.max(.1,Number(riskPct)||1)),capital=Math.max(100,Number(balance)||10000),qty=Math.floor(Math.min(capital*risk/100/distance,capital*.95/price)),quantity=canLong||canShort?qty:0;
  const flowBuy=aggTrades?.buyNotionalUsd??historicalBuy,flowSell=aggTrades?.sellNotionalUsd??historicalSell;
  return{symbol,timeframe,quote:{price,change24hPct:ticker.changePct??null,quoteVolume24h:ticker.quoteVolume??null,source:ticker.source||'Binance'},session:marketClock(),technical:{trend:trendUp?'صعودی':'نزولی',rsi:rsiValue,macdHistogram:macdHist,atr,atrPct:atr/price*100,support:structure.channelLow,resistance:structure.channelHigh,score:technicalScore,factors:technical,structure:{channelHigh:structure.channelHigh,channelLow:structure.channelLow,channelLongBreak:structure.channelLongBreak,channelShortBreak:structure.channelShortBreak,swingHigh:structure.swingHigh,swingLow:structure.swingLow,fibonacciLevels:structure.fibonacciLevels,fibLongBounce:structure.fibLongBounce,fibShortReject:structure.fibShortReject}},pressure:{takerImbalancePct:tradePressure,orderBookImbalancePct:bookPressure,historicalVolumePressurePct:historicalPressure,label:tradePressure>10?'غلبهٔ خریدار تهاجمی':tradePressure< -10?'غلبهٔ فروشنده تهاجمی':'تقریباً متعادل',buyNotionalUsd:flowBuy,sellNotionalUsd:flowSell,orderBookBidUsd:orderBook?.bidNotionalUsd??null,orderBookAskUsd:orderBook?.askNotionalUsd??null,largeBuyTrades:aggTrades?.largeBuyTrades||0,largeSellTrades:aggTrades?.largeSellTrades||0,veryLargeBuyTrades:aggTrades?.veryLargeBuyTrades||0,veryLargeSellTrades:aggTrades?.veryLargeSellTrades||0,note:'جریان taker و عمق دفتر سفارش از snapshot زنده؛ backtest فقط فشار تاریخی کندل را دارد.'},fundamentals:fundamentals?{...fundamentals,score:fundamentalScore,factors:fundamentalFactors,available:true}:{available:false,score:0,factors:[]},news:{items:newsItems,score:newsScore,available:newsAvailable,status:news?.status||'منبع خبر تنظیم نشده.'},decision:{direction,score:combined,confidencePct:Math.min(95,Math.round(42+Math.abs(combined)*7)),layers:[{label:'تکنیکال + کانال + فیبوناچی',score:technicalScore},{label:'فشار taker',score:flowScore},{label:'دفتر سفارش',score:bookScore},{label:'فاندامنتال/عرضه',score:fundamentalScore},{label:'خبر تازه',score:newsLayer}],confirmation:{long:longConditions,short:shortConditions},note:'ورود فقط وقتی پیشنهاد می‌شود که روند، شکست کانال یا واکنش فیبوناچی، فشار بازار، داده بنیادی و خبر تازه هم‌جهت باشند؛ امتیاز احتمال موفقیت نیست.'},plan:{referenceEntry:price,longStop:price-distance,longTarget:price+2*distance,shortStop:price+distance,shortTarget:price-2*distance,stopPct,targetPct:stopPct*2,riskReward:'۱:۲',riskPct:risk,riskAmount:quantity?quantity*distance:0,quantity,estimatedNotional:quantity*price,entryAllowed:canLong||canShort,entrySide:candidateSide,entryConditions:candidateSide==='long'?longConditions:shortConditions,note:quantity?'فقط Paper Trade/طرح مجازی؛ هیچ سفارش صرافی ارسال نمی‌شود.':'همهٔ تأییدهای چندعاملی هم‌زمان کامل نیست؛ ورود مجاز نیست.'},dataStatus:{fundamentals:fundamentals?'CoinGecko':'دردسترس نیست',news:newsAvailable?'خبر تازه':'دردسترس نیست',orderBook:orderBook?.source||'داده جاری در دسترس نیست',tradeFlow:aggTrades?.source||'تخمین کندلی'}};
}
module.exports={analyze,ema,rsi,marketStructure};
