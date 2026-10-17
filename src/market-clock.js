function marketClock(now = new Date()) {
 const zone=Intl.DateTimeFormat().resolvedOptions().timeZone||'local';
 const parts=new Intl.DateTimeFormat('fa-IR',{timeZone:zone,weekday:'long',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(now);
 const values=Object.fromEntries(parts.map(x=>[x.type,x.value]));
 return{marketOpen:true,market:'بازار رمزارزها · شبانه‌روزی',timeZone:zone,localTime:`${values.weekday} ${values.hour}:${values.minute}`,note:'کریپتو ۲۴/۷ معامله می‌شود؛ برنامه‌ریزی تحلیل بر اساس ساعت محلی دستگاه است.'};
}
module.exports={marketClock};
