function ema(values, period) { const out=Array(values.length).fill(null), k=2/(period+1); let state=null; values.forEach((v,i)=>{state=state===null?v:v*k+state*(1-k);if(i>=period-1)out[i]=state;});return out; }
function rsi(values, period=14) {const out=Array(values.length).fill(null);if(values.length<=period)return out;let gain=0,loss=0;for(let i=1;i<=period;i++){const d=values[i]-values[i-1];gain+=Math.max(0,d);loss+=Math.max(0,-d);}let ag=gain/period,al=loss/period;out[period]=al===0?100:100-100/(1+ag/al);for(let i=period+1;i<values.length;i++){const d=values[i]-values[i-1];ag=(ag*(period-1)+Math.max(0,d))/period;al=(al*(period-1)+Math.max(0,-d))/period;out[i]=al===0?100:100-100/(1+ag/al);}return out;}
function atrAt(candles,index,period=14){let sum=0,count=0;for(let i=Math.max(1,index-period+1);i<=index;i++){const p=candles[i-1].close;sum+=Math.max(candles[i].high-candles[i].low,Math.abs(candles[i].high-p),Math.abs(candles[i].low-p));count++;}return sum/Math.max(1,count);}
function factorFrame(candles){const closes=candles.map(c=>c.close), fast=ema(closes,12), slow=ema(closes,26);const macd=closes.map((_,i)=>fast[i]!=null&&slow[i]!=null?fast[i]-slow[i]:0),signal=ema(macd,9),osc=rsi(closes,14);return candles.map((c,i)=>{
  const range=Math.max(1e-12,c.high-c.low);let buy,sell;
  if(Number.isFinite(c.takerBuyQuoteVolume)&&Number.isFinite(c.quoteVolume)){buy=c.takerBuyQuoteVolume;sell=Math.max(0,c.quoteVolume-buy);}else{const buyShare=Math.max(0,Math.min(1,(c.close-c.low)/range));buy=(c.quoteVolume||c.volume||0)*buyShare;sell=(c.quoteVolume||c.volume||0)*(1-buyShare);}
  let flowBuy=0,flowSell=0;for(let j=Math.max(0,i-19);j<=i;j++){const cj=candles[j],r=Math.max(1e-12,cj.high-cj.low);if(Number.isFinite(cj.takerBuyQuoteVolume)&&Number.isFinite(cj.quoteVolume)){flowBuy+=cj.takerBuyQuoteVolume;flowSell+=Math.max(0,cj.quoteVolume-cj.takerBuyQuoteVolume);}else{const portion=Math.max(0,Math.min(1,(cj.close-cj.low)/r));const vol=cj.quoteVolume||cj.volume||0;flowBuy+=vol*portion;flowSell+=vol*(1-portion);}}
  const flowTotal=flowBuy+flowSell,flowPct=flowTotal?(flowBuy-flowSell)/flowTotal*100:0;
  let high=-Infinity,low=Infinity;for(let j=Math.max(0,i-20);j<i;j++){high=Math.max(high,candles[j].high);low=Math.min(low,candles[j].low);}
  let score=0;if(fast[i]!=null&&slow[i]!=null)score+=fast[i]>slow[i]?1:-1;if(signal[i]!=null)score+=macd[i]>=signal[i]?1:-1;if(osc[i]!=null)score+=osc[i]<35?1:osc[i]>70?-1:0;score+=flowPct>=15?1:flowPct<=-15?-1:0;if(c.close>high)score++;else if(c.close<low)score--;
  return{score,flowPct,rsi:osc[i],atr:atrAt(candles,i),emaFast:fast[i],emaSlow:slow[i]};
});}
function runBacktest({symbol,candles,riskPct=1,initialBalance=10000,feeRate=.001}){
 if(!Array.isArray(candles)||candles.length<80)throw new Error('برای بک‌تست معتبر حداقل ۸۰ کندل لازم است.');
 const start= Math.max(100,Number(initialBalance)||10000),risk=Math.min(3,Math.max(.1,Number(riskPct)||1))/100,fees=Math.min(.005,Math.max(0,Number(feeRate)||0));
 const frames=factorFrame(candles);let balance=start,peak=start,maxDd=0,position=null;const trades=[],equity=[];
 function close(i,price,reason){const raw=position.side*(price-position.entry)/position.entry*position.notional;const pnl=raw-position.notional*fees*2;balance+=pnl;trades.push({side:position.side>0?'لانگ':'شورت',entry:position.entry,exit:price,pnl,pnlPct:pnl/position.notional*100,entryTime:candles[position.index].time,exitTime:candles[i].time,reason,stopPct:position.stopPct*100,targetPct:position.targetPct*100});position=null;}
 for(let i=30;i<candles.length;i++){
   let closedAtProtectiveLevel=false;
   if(position){const stop=position.entry*(1-position.side*position.stopPct),target=position.entry*(1+position.side*position.targetPct);const stopHit=position.side>0?candles[i].low<=stop:candles[i].high>=stop;const targetHit=position.side>0?candles[i].high>=target:candles[i].low<=target;
     if(stopHit){const fill=position.side>0?Math.min(stop,candles[i].open):Math.max(stop,candles[i].open);close(i,fill,'حد ضرر ATR');closedAtProtectiveLevel=true;}
     else if(targetHit){close(i,target,'حد سود ۲R');closedAtProtectiveLevel=true;}
   }
   const signal=frames[i].score>=3?1:frames[i].score<=-3?-1:0;
   if(signal&&position&&signal!==position.side)close(i,candles[i].close,'چرخش امتیاز چندعاملی');
   if(signal&&!position&&!closedAtProtectiveLevel){const atr=frames[i].atr;const stopPct=Math.min(.08,Math.max(.004,atr/candles[i].close*1.5));const targetPct=stopPct*2;const notional=Math.min(balance*.95,balance*risk/stopPct);position={side:signal,entry:candles[i].close,index:i,notional,stopPct,targetPct};}
   const mark=position?balance+position.side*(candles[i].close-position.entry)/position.entry*position.notional:balance;equity.push(mark);peak=Math.max(peak,mark);maxDd=Math.max(maxDd,(peak-mark)/peak);
 }
 if(position)close(candles.length-1,candles.at(-1).close,'پایان تاریخچه');
 const winners=trades.filter(t=>t.pnl>0),losers=trades.filter(t=>t.pnl<=0);const grossWin=winners.reduce((s,t)=>s+t.pnl,0),grossLoss=Math.abs(losers.reduce((s,t)=>s+t.pnl,0));
 const pct=(balance/start-1)*100;
 return{symbol,summary:{initialBalance:start,finalBalance:balance,netProfit:balance-start,returnPct:pct,maxDrawdownPct:maxDd*100,winRatePct:trades.length?winners.length/trades.length*100:0,profitFactor:grossLoss?grossWin/grossLoss:grossWin?null:0,trades:trades.length,wins:winners.length,losses:losers.length,feePct:fees*100,riskPct:risk*100},trades:trades.slice(-12).reverse(),candlesCount:candles.length,from:candles[0].time,to:candles.at(-1).time,equity};
}
module.exports={runBacktest,ema,rsi,atrAt,factorFrame};
