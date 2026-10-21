# ربات انتشار ویدیو در پیج اینستاگرام

ویدیوها را از یک پوشه برمی‌دارد، کپشن هرکدام را پیدا می‌کند و آن‌ها را به‌صورت **ریلز** در پیج اینستاگرام منتشر می‌کند. یک‌باره (`post`) یا با رصد مداوم پوشه (`watch`).

نوشته‌شده با **API رسمی اینستاگرام (Instagram Platform / Graph API)**. هیچ لاگین با نام کاربری و رمز، کوکی یا اتوماسیون مرورگر در کار نیست؛ این کار خلاف قوانین اینستاگرام است و به بلاک شدن اکانت می‌انجامد.

```
instagram-uploader/
├── cli.js                 ← نقطهٔ ورود (check / list / post / watch / serve / status / reset)
├── .env.example           ← نمونهٔ تنظیمات
├── START-WINDOWS.md       ← راهنمای سادهٔ قدم‌به‌قدم ویندوز
├── setup.bat              ← تنظیم یک‌بارهٔ ویندوز (می‌پرسد و ذخیره می‌کند)
├── media/                 ← ویدیوها را اینجا بریز (+ کپشن‌ها)
│   └── posted/            ← اگر آرشیو فعال باشد، فایل‌های منتشرشده اینجا می‌روند
├── data/                  ← حافظهٔ ربات و لاگ قفل (خودکار ساخته می‌شود، commit نشود)
├── src/
│   ├── config.js          ← خواندن .env و اعتبارسنجی
│   ├── graph.js           ← ساخت container، انتظار پردازش، انتشار، سهمیه، تمدید توکن
│   ├── media-server.js    ← سرور کوچک فایل ویدیو با URL غیرقابل‌حدس + Range
│   ├── captions.js        ← انتخاب و آماده‌سازی کپشن
│   ├── library.js         ← پیمایش پوشه و حافظهٔ «چه چیزی منتشر شده»
│   ├── probe.js           ← بررسی اختیاری ویدیو با ffprobe
│   └── publisher.js       ← جریان کامل یک اجرا
└── test/uploader.test.js  ← تست‌های آفلاین
```

> 🪟 **کاربر ویندوز و تازه‌کار؟** فایل [START-WINDOWS.md](START-WINDOWS.md) را باز کن؛
> همان مسیر با زبان ساده و بدون پیش‌فرض فنی نوشته شده. بعد از گرفتن توکن، فقط
> `setup.bat` را دوبار کلیک کن؛ خودش سه مقدار را می‌پرسد، در `.env` می‌نویسد و تست می‌کند.

## ۰) پیش‌نیازها

1. **Node.js 18 یا جدیدتر** (برای `fetch` داخلی). بدون هیچ پکیج npm کار می‌کند.
2. **اکانت حرفه‌ای اینستاگرام**: Business یا Creator. اکانت شخصی نمی‌تواند از API منتشر کند.
3. **یک اپ متا** با محصول Instagram و مجوز `instagram_business_content_publish` (و `instagram_business_basic`).
4. **ffprobe** (اختیاری ولی توصیه‌شده) برای بررسی مدت/کدک: بخشی از FFmpeg است.
5. **یک آدرس عمومی HTTPS** برای فایل ویدیو — بخش ۳ را ببین.

## ۱) گرفتن شناسه و توکن

اینستاگرام فایل را از تو **آپلود نمی‌گیرد**؛ فقط یک `video_url` عمومی می‌پذیرد. پس توکن و شناسه لازم است:

1. به [developers.facebook.com](https://developers.facebook.com/apps) برو و یک App از نوع **Business** بساز.
2. محصول **Instagram** را به اپ اضافه کن. یکی از دو مسیر:
   - **Instagram API with Instagram Login** — بدون نیاز به صفحهٔ فیسبوک. آدرس API: `https://graph.instagram.com`.
     در تنظیمات `IG_GRAPH_BASE_URL=https://graph.instagram.com` بگذار.
   - **Instagram API with Facebook Login** — اکانت اینستاگرام باید به یک صفحهٔ فیسبوک وصل باشد. آدرس API: `https://graph.facebook.com` (پیش‌فرض).
3. دسترسی‌ها را بده: `instagram_business_basic` و `instagram_business_content_publish`.
   در حالت توسعه، اکانت خودت به‌عنوان Tester/Admin کافی است؛ برای اکانت‌های دیگر باید App Review بگیری.
4. با **Graph API Explorer** یا Business Login یک توکن بگیر، سپس توکن کوتاه‌مدت را به بلندمدت تبدیل کن:
   ```bash
   curl "https://graph.instagram.com/access_token?grant_type=ig_exchange_token&client_secret=APP_SECRET&access_token=SHORT_TOKEN"
   ```
   توکن بلندمدت **۶۰ روز** اعتبار دارد. با `refresh-token --save` هر وقت خواستی ۶۰ روز دیگر تمدیدش کن.
   (شرط تمدید: توکن حداقل ۲۴ ساعت عمر داشته و منقضی نشده باشد.)
5. شناسهٔ اکانت (`IG_USER_ID`) را بگیر — یک عدد بلند، نه `@username`:
   ```bash
   curl "https://graph.facebook.com/v23.0/me/accounts?access_token=TOKEN"
   # سپس برای هر صفحه:
   curl "https://graph.facebook.com/v23.0/PAGE_ID?fields=instagram_business_account&access_token=TOKEN"
   ```
   در مسیر Instagram Login: `curl "https://graph.instagram.com/v23.0/me?fields=user_id,username&access_token=TOKEN"`.

> توکن را در چت، اسکرین‌شات یا GitHub نگذار. فقط در `instagram-uploader/.env` (که git-ignore است) یا متغیر محیطی.

## ۲) تنظیمات

```bash
cd instagram-uploader
cp .env.example .env      # در ویندوز: copy .env.example .env
```

حداقل این‌ها را پر کن:

| متغیر | معنی |
| --- | --- |
| `IG_USER_ID` | شناسهٔ اکانت حرفه‌ای |
| `IG_ACCESS_TOKEN` | توکن بلندمدت |
| `IG_PUBLIC_BASE_URL` | آدرس عمومی HTTPS که به سرور ویدیوی محلی وصل است (بخش ۳) |
| `IG_MEDIA_DIR` | پوشهٔ ویدیوها، پیش‌فرض `media` |
| `IG_DEFAULT_CAPTION` | کپشن پیش‌فرض وقتی هیچ متنی پیدا نشود؛ `{name}` = نام فایل |
| `IG_DEFAULT_HASHTAGS` | هشتگ‌های پیش‌فرض (اگر کپشن خودش هشتگ نداشته باشد) |

بقیه: `IG_POST_INTERVAL_SEC` (فاصلهٔ بین دو پست، پیش‌فرض ۴۵ ثانیه)، `IG_MAX_POSTS_PER_RUN`، `IG_DAILY_LIMIT` (سقف ایمنی محلی)، `IG_WATCH_INTERVAL_SEC`، `IG_ARCHIVE_AFTER_POST`، `IG_DRY_RUN`.

## ۳) آدرس عمومی برای ویدیو (مهم‌ترین بخش)

اینستاگرام سرور را از بیرون خودش صدا می‌زند، پس `localhost` کار نمی‌کند. ربات یک سرور کوچک روی `IG_MEDIA_PORT` (پیش‌فرض ۸۷۸۷) بالا می‌آورد که فقط فایل‌های ویدیویی را با یک رشتهٔ مخفی در مسیر سرو می‌کند، سپس یک تونل جلوی آن می‌گذاری:

```bash
# ترمینال ۱
npx cloudflared tunnel --url http://localhost:8787
# → آدرس شبیه https://random-words.trycloudflare.com

# یا با ngrok
ngrok http 8787
```

آدرس را در `.env` بگذار و بعد ربات را اجرا کن:

```
IG_PUBLIC_BASE_URL=https://random-words.trycloudflare.com
```

سپس صحت همه‌چیز را چک کن:

```bash
node cli.js check
```

`check` توکن و نوع اکانت را می‌پرسد، سهمیهٔ انتشار را نشان می‌دهد و تست می‌کند که اولین ویدیو از بیرون با `https` در دسترس است. نکته: آدرس‌های `trycloudflare.com` موقتی‌اند؛ برای ربات همیشه‌روشن یک **Named Tunnel** کلادفلر یا دامنهٔ خودت را روی سرور بگذار و `IG_PUBLIC_BASE_URL` را ثابت کن.

اگر ربات و سرور روی یک ماشین همیشه‌روشن اجرا می‌شوند، می‌توانی سرور فایل را جدا نگه داری:

```bash
node cli.js serve     # فقط سرور فایل؛ تونل و ربات جای دیگر
```

## ۴) کپشن‌ها

برای هر ویدیو، متن به این ترتیب انتخاب می‌شود:

1. `media/captions.json` — دقیق‌ترین راه. سه شکل مقدار:
   ```json
   {
     "_default": "کپشن همهٔ ویدیوهای بدون متن",
     "clip1.mp4": "متن ساده",
     "clip2.mp4": { "caption": "متن + کاور", "cover": "https://…/cover.jpg", "thumb_offset": 1500 }
   }
   ```
   (کلید می‌تواند نام کامل فایل یا نام بدون پسوند باشد؛ فایل نمونه: `media/captions.example.json`.)
2. فایل متنی کنار ویدیو: `clip1.txt` یا `clip1.caption.txt`.
3. `IG_DEFAULT_CAPTION` با پلیس‌هولدرهای `{name}`، `{file}`، `{path}`، `{index}`، `{date}`.

بعد از انتخاب، اگر کپشن هیچ هشتگی نداشت `IG_DEFAULT_HASHTAGS` به آن اضافه می‌شود (حداکثر ۳۰ هشتگ)، طول متن به ۲۲۰۰ کاراکتر محدود می‌شود و `\r\n`ها اصلاح می‌شوند.

## ۵) اجرا

```bash
cd instagram-uploader

node cli.js list                 # چه ویدیوهایی هست؟ چه چیزی منتشر شده؟ کپشن هرکدام چه می‌شود؟
node cli.js post --dry-run       # شبیه‌سازی کامل، بدون تماس با اینستاگرام
node cli.js post                 # انتشار همهٔ ویدیوهای جدید (تا سقف اجرا)
node cli.js post --limit 1       # فقط یکی
node cli.js post --file clip2.mp4
node cli.js watch                # رصد مداوم پوشه؛ هر ویدیوی تازه خودکار منتشر می‌شود
node cli.js watch --interval 600 --archive
node cli.js status               # آمار، سهمیهٔ رسمی اینستاگرام، آخرین خطاها
node cli.js refresh-token --save # تمدید توکن ۶۰ روزه
node cli.js reset --file clip1.mp4   # فراموش کردن یک فایل تا دوباره ارسال شود
```

از داخل ریشهٔ پروژه هم می‌شود: `npm run ig:post`، `npm run ig:watch` و… .

هر اجرا فقط **ویدیوهای جدید** را می‌فرستد؛ حافظهٔ آن در `data/state.json` است، پس اجرای دوباره چیزی را تکرار نمی‌کند.

### زمان‌بندی خودکار

**لینوکس/macOS (cron)** — هر ۱۵ دقیقه یک‌بار، هر ویدیوی تازه منتشر می‌شود:

```cron
*/15 * * * * cd /path/to/instagram-uploader && /usr/bin/node cli.js post --limit 1 >> data/cron.log 2>&1
```

**ویندوز (Task Scheduler)** — یک Task بساز، برنامه: `node.exe`، آرگومان‌ها: `cli.js post --limit 1`، «Start in»: مسیر پوشهٔ `instagram-uploader`.

**همیشه‌روشن با pm2:**

```bash
pm2 start cli.js --name ig-uploader -- watch --archive
pm2 save
```

اگر از cron استفاده می‌کنی، سرور فایل باید همیشه بالا باشد (`node cli.js serve` زیر pm2 یا systemd) وگرنه اینستاگرام موقع پردازش ویدیو به فایل نمی‌رسد.

## ۶) اینستاگرام چه چیزی می‌پذیرد

- فقط **ویدیو برای ریلز**: `MP4` یا `MOV`، کدک H.264 + صدای AAC، عمودی **۹:۱۶** (۱۰۸۰×۱۹۲۰)، بین **۳ ثانیه تا ۱۵ دقیقه**.
- سقف حجم: منابع متا بین ۳۰۰ مگابایت و ۱ گیگابایت اختلاف دارند؛ برای اطمینان **زیر ۳۰۰ مگابایت** بمان. (این عدد در `src/probe.js` قابل تغییر است.)
- کپشن حداکثر ۲۲۰۰ کاراکتر، حداکثر ۳۰ هشتگ.
- container ساخته‌شده اگر منتشر نشود بعد از **۲۴ ساعت** منقضی می‌شود؛ ربات هر بار از نو می‌سازد.
- سقف انتشار **در هر ۲۴ ساعت شناور** است و متا خودش عدد را چند بار عوض کرده. عدد واقعی را با `status` یا `check` از `GET /{ig-id}/content_publishing_limit` بگیر (معمولاً ۵۰ یا ۱۰۰). `IG_DAILY_LIMIT` فقط سقف ایمنی سمت خودت است.
- برای هر ویدیو: ساخت container → انتظار تا `FINISHED` (چند ثانیه تا چند دقیقه) → `media_publish`. ربات وضعیت را هر ۱۰ ثانیه می‌پرسد و تا ۲۰ دقیقه صبر می‌کند.

## ۷) خطاهای رایج

| نشانه | معنی و راه‌حل |
| --- | --- |
| `Invalid OAuth access token` (کد ۱۹۰) | توکن منقضی/غلط است → توکن تازه بگیر یا `refresh-token --save`. |
| `(#9004)` یا `Only photo or video can be accepted` | فایل یا URL مشکل دارد: فرمت/کدک، یا اینستاگرام فایل را دریافت نکرده. با `check` و مرورگر امتحان کن. |
| `Media download has failed` / `timeout` | اینستاگرام به `IG_PUBLIC_BASE_URL` نرسیده یا پاسخ کند بوده. تونل، فایروال و `https` را چک کن. |
| کد ۹ / زیرکد `2207042` | سقف انتشار ۲۴ ساعته پر شده → صبر کن یا `status` را ببین. |
| `status_code: ERROR` در پردازش | ویدیو برای اینستاگرام قابل پردازش نیست؛ خروجی ffprobe را ببین و با H.264/AAC دوباره انکود کن. |
| `اکانت BUSINESS/CREATOR نیست` | اکانت را حرفه‌ای کن و (در مسیر Facebook Login) به یک صفحهٔ فیسبوک وصل کن. |
| `یک اجرای دیگر در حال کار است` | فایل `data/uploader.lock` از یک کرش باقی مانده؛ اگر مطمئنی اجرایی در جریان نیست پاکش کن. |

## ۸) امنیت و نگهداری

- `.env`، `data/` و فایل‌های ویدیویی در `.gitignore` هستند؛ توکن و محتوا commit نمی‌شوند.
- سرور فایل هر فایل را با رشتهٔ مخفیِ مسیر سرو می‌کند؛ ولی هر کسی که URL را داشته باشد تا وقتی پست منتشر نشده می‌تواند دانلودش کند. رشته را با `IG_MEDIA_SECRET` عوض کن (وگرنه خودکار ساخته می‌شود).
- سرور فقط پسوندهای ویدیویی را می‌دهد، از `..` جلوگیری می‌کند و پوشهٔ آرشیو را هم سرو می‌کند (اگر نمی‌خواهی، آن را بیرون از `media/` بگذار).
- بعد از انتشار موفق، اگر `IG_ARCHIVE_AFTER_POST=true` باشد فایل به `media/posted/` منتقل می‌شود؛ پوشهٔ آرشیو در پیمایش نادیده گرفته می‌شود.
- حافظهٔ ربات (`data/state.json`) با نام نسبی فایل + media id + زمان کار می‌کند؛ اگر فایل را تغییر نام دهی، دوباره ارسال می‌شود.

## ۹) تست

```bash
node --test instagram-uploader/test/*.test.js
```

تست‌ها آفلاین‌اند (Graph API با `fetch` جعلی شبیه‌سازی می‌شود) و سرور فایل، Range، مسیر امن، اولویت کپشن، حافظه، آرشیو، شکست/سهمیه و ساخت container را پوشش می‌دهند.

## محدودیت صریح

این ماژول فقط ریلز منتشر می‌کند (استوری و کاروسل پیاده‌سازی نشده) و ویدیو را بهینه/ترنسکود نمی‌کند. اگر ویدیو ۹:۱۶ و H.264 نباشد، اینستاگرام ممکن است آن را برش بدهد یا رد کند؛ `checkReelConstraints` فقط هشدار می‌دهد و جلوی ارسال را نمی‌گیرد.
