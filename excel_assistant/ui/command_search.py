"""
Command Search Registry & Matching Engine for MUYUM Excel Assistant.
Provides fast, fuzzy, and semantic Persian search across all Excel assistant features.
"""
from typing import List, Dict, Any, Optional

COMMAND_SEARCH_REGISTRY: List[Dict[str, Any]] = [
    # 💼 ۲. قابلیت‌های حسابداری
    {
        "title": "💵 محاسبه سود سالانه",
        "cmd": "محاسبه سود سالانه",
        "category": "حسابداری",
        "keywords": ["سود", "سود سالانه", "سالانه", "سال", "درآمد سالانه", "profit", "annual", "yearly", "محاسبه سود"]
    },
    {
        "title": "💵 محاسبه سود ماهانه",
        "cmd": "محاسبه سود ماهانه",
        "category": "حسابداری",
        "keywords": ["سود", "سود ماهانه", "ماهانه", "سود ماه", "درآمد ماهانه", "profit", "monthly", "محاسبه سود"]
    },
    {
        "title": "💵 محاسبه سود روزانه",
        "cmd": "محاسبه سود روزانه",
        "category": "حسابداری",
        "keywords": ["سود", "سود روزانه", "روزانه", "سود روز", "profit", "daily", "محاسبه سود"]
    },
    {
        "title": "💵 محاسبه سود کل دوره",
        "cmd": "محاسبه سود کل",
        "category": "حسابداری",
        "keywords": ["سود", "سود کل", "کل سود", "مجموع سود", "profit", "total", "محاسبه سود"]
    },
    {
        "title": "💵 محاسبه سود به روش فایفو (FIFO)",
        "cmd": "محاسبه سود به روش فایفو",
        "category": "حسابداری",
        "keywords": ["سود", "فایفو", "fifo", "اولین وارده", "بهای تمام شده", "cogs", "محاسبه سود"]
    },
    {
        "title": "💵 محاسبه سود به روش لایفو (LIFO)",
        "cmd": "محاسبه سود به روش لایفو",
        "category": "حسابداری",
        "keywords": ["سود", "لایفو", "lifo", "آخرین وارده", "تورم", "مالیات", "محاسبه سود"]
    },
    {
        "title": "💵 محاسبه سود با میانگین موزون (Average)",
        "cmd": "محاسبه سود با میانگین موزون",
        "category": "حسابداری",
        "keywords": ["سود", "میانگین", "موزون", "average", "محاسبه سود"]
    },
    {
        "title": "⚖️ مقایسه تطبیقی سود (فایفو vs لایفو vs میانگین)",
        "cmd": "مقایسه سود به روش فایفو و لایفو",
        "category": "حسابداری",
        "keywords": ["سود", "مقایسه", "فایفو", "لایفو", "تطبیقی", "compare", "محاسبه سود"]
    },
    {
        "title": "🏷️ انتخاب روش محاسبه سود (فایفو/لایفو/میانگین)",
        "cmd": "روش سود",
        "category": "حسابداری",
        "keywords": ["سود", "روش سود", "انتخاب سود", "ویزارد سود", "محاسبه سود"]
    },
    {
        "title": "⚙️ تنظیمات دستی ستون‌های محاسبه سود",
        "cmd": "تنظیمات سود",
        "category": "حسابداری",
        "keywords": ["سود", "تنظیمات سود", "دستی ستون سود", "محاسبه سود"]
    },
    {
        "title": "🏷️ انتخاب روش محاسبه بهای تمام شده (فایفو/لایفو/میانگین)",
        "cmd": "روش بهای تمام شده",
        "category": "حسابداری",
        "keywords": ["بهای تمام شده", "روش بهای تمام شده", "انتخاب بهای تمام شده", "ویزارد بهای تمام شده", "cogs", "فایفو", "لایفو", "قیمت تمام شده"]
    },
    {
        "title": "🏷️ محاسبه بهای تمام شده کالای فروش رفته (COGS)",
        "cmd": "محاسبه بهای تمام شده کالای فروش رفته",
        "category": "حسابداری",
        "keywords": ["بهای تمام شده", "بهای تمام‌شده", "قیمت تمام شده", "قیمت تمام‌شده", "cogs", "کالای فروش رفته", "صورت بهای تمام شده", "جدول بهای تمام شده", "تمام شده"]
    },
    {
        "title": "🏷️ بهای تمام شده به روش فایفو (FIFO)",
        "cmd": "محاسبه بهای تمام شده به روش فایفو",
        "category": "حسابداری",
        "keywords": ["بهای تمام شده", "فایفو", "fifo", "cogs", "اولین وارده", "قیمت تمام شده"]
    },
    {
        "title": "🏷️ بهای تمام شده به روش لایفو (LIFO)",
        "cmd": "محاسبه بهای تمام شده به روش لایفو",
        "category": "حسابداری",
        "keywords": ["بهای تمام شده", "لایفو", "lifo", "cogs", "آخرین وارده", "تورم", "قیمت تمام شده"]
    },
    {
        "title": "⚖️ مقایسه بهای تمام شده (فایفو vs لایفو vs میانگین)",
        "cmd": "مقایسه بهای تمام شده فایفو و لایفو",
        "category": "حسابداری",
        "keywords": ["بهای تمام شده", "مقایسه بهای تمام شده", "فایفو", "لایفو", "میانگین", "cogs", "تطبیقی"]
    },
    {
        "title": "⏱️ محاسبه ساعت کاری، ورود و خروج کارمندان (تایم‌شیت)",
        "cmd": "محاسبه ساعت کاری",
        "category": "حسابداری",
        "keywords": [
            "ساعت کاری", "ساعت کار", "ساعت کارکرد", "ساعات کاری", "ورود و خروج", "ورود خروج",
            "حضور و غیاب", "حضور غیاب", "تایم شیت", "تایمشیت", "اضافه کار", "اضافه کاری",
            "کسری کار", "مرخصی", "تردد", "پرسنل", "کارمندان", "تایم‌شیت"
        ]
    },
    {
        "title": "💼 لیست حقوق و دستمزد، بیمه و مالیات",
        "cmd": "محاسبه حقوق و دستمزد",
        "category": "حسابداری",
        "keywords": ["حقوق", "دستمزد", "فیش حقوقی", "فیش حقوق", "بیمه", "مالیات", "payroll", "salary"]
    },
    {
        "title": "💰 گزارش جامع فروش (تشخیص خودکار ستون‌ها)",
        "cmd": "محاسبه فروش",
        "category": "حسابداری",
        "keywords": ["فروش", "فروش کل", "فروش ماهانه", "فروش سالانه", "فروش روزانه", "sales", "گزارش فروش"]
    },
    {
        "title": "⚖️ مغایرت‌گیری دو ستون (Reconciliation)",
        "cmd": "مغایرت‌گیری ستون A و B",
        "category": "حسابداری",
        "keywords": ["مغایرت", "مغایرت گیری", "اختلاف", "تطبیق", "reconciliation"]
    },
    {
        "title": "📊 تهیه تراز آزمایشی چهار ستونه",
        "cmd": "تهیه تراز آزمایشی",
        "category": "حسابداری",
        "keywords": ["تراز", "تراز آزمایشی", "چهار ستونه", "trial balance"]
    },
    {
        "title": "💳 محاسبه مانده حساب جاری (بدهکار/بستانکار)",
        "cmd": "محاسبه مانده حساب",
        "category": "حسابداری",
        "keywords": ["مانده", "مانده حساب", "بدهکار", "بستانکار", "تسویه"]
    },
    {
        "title": "📉 محاسبه جدول استهلاک دارایی‌ها",
        "cmd": "محاسبه استهلاک",
        "category": "حسابداری",
        "keywords": ["استهلاک", "دارایی", "استهلاک انباشته", "depreciation"]
    },
    {
        "title": "📈 تهیه صورت سود و زیان جامع دوره",
        "cmd": "صورت سود و زیان جامع",
        "category": "حسابداری",
        "keywords": ["صورت سود و زیان", "سود و زیان", "سود", "زیان", "income statement", "بیلان"]
    },
    {
        "title": "🏛️ تهیه ترازنامه مالی (بیلان)",
        "cmd": "ترازنامه",
        "category": "حسابداری",
        "keywords": ["ترازنامه", "صورت وضعیت مالی", "بیلان", "balance sheet"]
    },
    {
        "title": "🏷️ محاسبه تخفیف کالاها",
        "cmd": "محاسبه تخفیف",
        "category": "حسابداری",
        "keywords": ["تخفیف", "درصد تخفیف", "discount", "off"]
    },

    # 🛒 ۳. قابلیت‌های فروش و انبارداری
    {
        "title": "🧾 صدور و ثبت فاکتور رسمی فروش",
        "cmd": "ثبت فاکتور فروش",
        "category": "فروش و انبارداری",
        "keywords": ["فاکتور", "فاکتور فروش", "صدور فاکتور", "invoice"]
    },
    {
        "title": "🛒 صدور و ثبت فاکتور رسمی خرید",
        "cmd": "ثبت فاکتور خرید",
        "category": "فروش و انبارداری",
        "keywords": ["خرید", "فاکتور خرید", "purchase"]
    },
    {
        "title": "📦 کنترل موجودی و هشدار کمبود انبار",
        "cmd": "موجودی انبار و هشدار کمبود موجودی",
        "category": "فروش و انبارداری",
        "keywords": ["موجودی", "انبار", "کمبود انبار", "کسری موجودی", "هشدار انبار"]
    },
    {
        "title": "📊 ارزش‌گذاری موجودی انبار به روش فایفو (FIFO)",
        "cmd": "ارزش‌گذاری انبار به روش فایفو",
        "category": "فروش و انبارداری",
        "keywords": ["انبار", "ارزش گذاری", "فایفو", "موجودی انبار"]
    },
    {
        "title": "📊 ارزش‌گذاری موجودی انبار به روش لایفو (LIFO)",
        "cmd": "ارزش‌گذاری انبار به روش لایفو",
        "category": "فروش و انبارداری",
        "keywords": ["انبار", "ارزش گذاری", "لایفو"]
    },
    {
        "title": "📊 ارزش‌گذاری موجودی انبار با میانگین موزون",
        "cmd": "ارزش‌گذاری انبار به روش میانگین موزون",
        "category": "فروش و انبارداری",
        "keywords": ["انبار", "ارزش گذاری", "میانگین موزون", "کاردکس"]
    },
    {
        "title": "💰 محاسبه ارزش ریالی کل موجودی انبار",
        "cmd": "ارزش کل موجودی انبار",
        "category": "فروش و انبارداری",
        "keywords": ["ارزش انبار", "ارزش ریالی", "موجودی کل"]
    },
    {
        "title": "👥 محاسبه بدهی و طلب مشتریان",
        "cmd": "محاسبه بدهی و طلب مشتریان",
        "category": "فروش و انبارداری",
        "keywords": ["مشتری", "بدهی", "طلب", "حساب مشتری"]
    },
    {
        "title": "🏆 گزارش پرفروش‌ترین کالاها",
        "cmd": "گزارش پرفروش‌ترین کالاها",
        "category": "فروش و انبارداری",
        "keywords": ["پرفروش", "کالای پرفروش", "پرفروش ترین", "top selling"]
    },
    {
        "title": "🏷️ افزایش درصدی قیمت‌های کالا",
        "cmd": "قیمت‌های ستون C را ۱۵ درصد افزایش بده",
        "category": "فروش و انبارداری",
        "keywords": ["افزایش قیمت", "درصد افزایش", "تغییر قیمت"]
    },

    # ⚡ ۱. قابلیت‌های ضروری
    {
        "title": "🔍 فیلتر هوشمند اطلاعات جدول",
        "cmd": "فیلتر هوشمند جدول",
        "category": "ضروری",
        "keywords": ["فیلتر", "جستجو", "filter", "autofilter"]
    },
    {
        "title": "🔄 مرتب‌سازی صعودی داده‌ها",
        "cmd": "مرتب‌سازی صعودی ستون A",
        "category": "ضروری",
        "keywords": ["مرتب سازی", "صعودی", "الفبا", "sort"]
    },
    {
        "title": "🔽 مرتب‌سازی نزولی داده‌ها",
        "cmd": "مرتب‌سازی نزولی ستون A",
        "category": "ضروری",
        "keywords": ["مرتب سازی", "نزولی", "sort desc"]
    },
    {
        "title": "🗑️ حذف موارد تکراری",
        "cmd": "حذف داده‌های تکراری",
        "category": "ضروری",
        "keywords": ["حذف تکراری", "تکراری", "duplicate", "remove duplicates"]
    },
    {
        "title": "🧹 حذف سطرهای خالی شیت",
        "cmd": "ردیف‌های خالی را حذف کن",
        "category": "ضروری",
        "keywords": ["حذف خالی", "ردیف خالی", "سطر خالی"]
    },
    {
        "title": "➕ جمع خودکار (AutoSum)",
        "cmd": "جمع خودکار ستون A",
        "category": "ضروری",
        "keywords": ["جمع", "مجموع", "جمع خودکار", "autosum", "sum"]
    },
    {
        "title": "📊 محاسبه میانگین داده‌ها",
        "cmd": "میانگین ستون A",
        "category": "ضروری",
        "keywords": ["میانگین", "معدل", "average"]
    },
    {
        "title": "🔺 بیشترین و کمترین مقدار (MAX / MIN)",
        "cmd": "بیشترین مقدار ستون A",
        "category": "ضروری",
        "keywords": ["بیشترین", "کمترین", "حداکثر", "حداقل", "max", "min"]
    },
    {
        "title": "🔢 شمارش اطلاعات و سلول‌های پر",
        "cmd": "شمارش سلول‌های پر ستون A",
        "category": "ضروری",
        "keywords": ["شمارش", "تعداد", "count"]
    },
    {
        "title": "🔢 تبدیل متن به عدد (اصلاح داده‌ها)",
        "cmd": "تبدیل متن به عدد در ستون A",
        "category": "ضروری",
        "keywords": ["تبدیل به عدد", "اصلاح متن", "text to number"]
    },
    {
        "title": "🧮 ۴ عمل اصلی ستون‌ها (جمع، تفریق، ضرب، تقسیم)",
        "cmd": "۴ عمل اصلی",
        "category": "ضروری",
        "keywords": ["چهار عمل اصلی", "جمع ستون", "تفریق", "ضرب", "تقسیم", "arithmetic"]
    },
    {
        "title": "📅 درج تاریخ روز در شیت",
        "cmd": "درج تاریخ در شیت",
        "category": "ضروری",
        "keywords": ["تاریخ", "درج تاریخ", "تاریخ شمسی", "امروز", "date"]
    },

    # 🤖 ۴. قابلیت‌های هوشمند و کاربردی
    {
        "title": "🔒 قفل یکتاسازی ستون (جلوگیری از ثبت تکراری)",
        "cmd": "قفل یکتایی ستون",
        "category": "هوشمند",
        "keywords": ["قفل یکتایی", "یکتا", "جلوگیری از تکراری", "unique validation"]
    },
    {
        "title": "🔓 لغو و برداشتن قفل یکتایی ستون",
        "cmd": "قفل یکتایی ستون را بردار",
        "category": "هوشمند",
        "keywords": ["برداشتن قفل", "لغو یکتایی", "unlock unique"]
    },
    {
        "title": "📊 تحلیل هوشمند شیت و داده‌ها",
        "cmd": "تحلیل ساختار فایل و شیت",
        "category": "هوشمند",
        "keywords": ["تحلیل", "آنالیز", "سلامت فایل", "analyze", "audit"]
    },
    {
        "title": "🩺 چکاپ عمیق سلامت و خطاهای فایل اکسل",
        "cmd": "سلامت فایل و بررسی خطاهای اکسل",
        "category": "هوشمند",
        "keywords": ["سلامت فایل", "خطا", "چکاپ", "audit"]
    },
    {
        "title": "📈 رسم خودکار نمودار گرافیکی",
        "cmd": "رسم نمودار ستونی",
        "category": "هوشمند",
        "keywords": ["نمودار", "چارت", "رسم چارت", "chart", "graph"]
    },
    {
        "title": "📋 تشخیص و ساخت جدول اکسل",
        "cmd": "یک جدول بساز",
        "category": "هوشمند",
        "keywords": ["جدول", "ساخت جدول", "create table"]
    },
    {
        "title": "✂️ تفکیک نام و نام خانوادگی",
        "cmd": "تفکیک نام و نام خانوادگی در ستون A",
        "category": "هوشمند",
        "keywords": ["تفکیک نام", "جداسازی نام", "split name"]
    },
    {
        "title": "🎨 رنگ‌آمیزی سلول‌های تکراری",
        "cmd": "داده‌های متنی تکراری را رنگ کن",
        "category": "هوشمند",
        "keywords": ["رنگ تکراری", "هایلایت تکراری", "highlight duplicates"]
    },
    {
        "title": "💲 تغییر فرمت به واحد پول (تومان)",
        "cmd": "فرمت ستون B را به تومان تغییر بده",
        "category": "هوشمند",
        "keywords": ["واحد پول", "فرمت پول", "تومان", "currency"]
    },
    {
        "title": "📈 تغییر فرمت به درصد (%)",
        "cmd": "فرمت ستون C را درصدی کن",
        "category": "هوشمند",
        "keywords": ["درصد", "فرمت درصد", "percentage"]
    },

    # ⚙️ ۵. امکانات عمومی و مدیریتی
    {
        "title": "💾 ذخیره تغییرات فایل اکسل (Ctrl+S)",
        "cmd": "ذخیره فایل",
        "category": "عمومی",
        "keywords": ["ذخیره", "سیو", "save"]
    },
    {
        "title": "💾 ذخیره نسخه پشتیبان (Backup)",
        "cmd": "ذخیره نسخه پشتیبان",
        "category": "عمومی",
        "keywords": ["پشتیبان", "بکاپ", "backup"]
    },
    {
        "title": "📄 تولید خروجی PDF استاندارد",
        "cmd": "خروجی PDF از شیت",
        "category": "عمومی",
        "keywords": ["پی دی اف", "pdf", "خروجی pdf", "پرینت"]
    },
    {
        "title": "↩️ لغو آخرین عملیات (Undo چندمرحله‌ای)",
        "cmd": "لغو آخرین عملیات",
        "category": "عمومی",
        "keywords": ["لغو", "آندو", "undo", "برگشت"]
    },
    {
        "title": "📐 تنظیم خودکار عرض ستون‌ها (AutoFit)",
        "cmd": "عرض ستون‌ها را خودکار تنظیم کن",
        "category": "عمومی",
        "keywords": ["تنظیم عرض", "عرض ستون", "autofit"]
    }
]


def normalize_persian_search_term(term: str) -> str:
    """Normalizes Arabic/Persian letters and whitespace for robust search matching."""
    if not term:
        return ""
    t = term.strip().lower()
    t = t.replace("ي", "ی").replace("ك", "ک").replace("آ", "ا").replace("أ", "ا").replace("إ", "ا")
    t = t.replace("\u200c", " ")  # zero-width non-joiner to space
    t = " ".join(t.split())
    return t


def search_commands(query: str, limit: int = 12) -> List[Dict[str, Any]]:
    """
    Intelligent fuzzy & semantic search across all assistant capabilities.
    Matches queries against titles, commands, and semantic keywords.
    """
    if not query or not query.strip():
        return []

    q_norm = normalize_persian_search_term(query)
    q_tokens = [tok for tok in q_norm.split() if len(tok) >= 1]
    if not q_tokens:
        return []

    results = []

    for idx, item in enumerate(COMMAND_SEARCH_REGISTRY):
        score = 0
        t_norm = normalize_persian_search_term(item["title"])
        c_norm = normalize_persian_search_term(item["cmd"])

        # 1. Exact match
        if q_norm == c_norm or q_norm == t_norm:
            score += 150
        # 2. Command starts with query or title starts with query
        elif c_norm.startswith(q_norm) or t_norm.startswith(q_norm):
            score += 100
        # 3. Query is inside command or title
        elif q_norm in c_norm:
            score += 70
        elif q_norm in t_norm:
            score += 60

        if item["cmd"].startswith("محاسبه "):
            score += 15

        # 4. Keyword matches (take max score among keyword hits)
        best_kw = 0
        for kw in item.get("keywords", []):
            kw_norm = normalize_persian_search_term(kw)
            if q_norm == kw_norm:
                best_kw = max(best_kw, 85)
            elif q_norm in kw_norm:
                best_kw = max(best_kw, 55)
            elif kw_norm in q_norm:
                best_kw = max(best_kw, 45)
        score += best_kw

        # 5. Token matches
        token_hits = 0
        for tok in q_tokens:
            if tok in c_norm.split():
                score += 50
                token_hits += 1
            elif tok in t_norm.split():
                score += 45
                token_hits += 1
            elif any(tok == normalize_persian_search_term(kw) for kw in item.get("keywords", [])):
                score += 40
                token_hits += 1
            elif tok in c_norm:
                score += 20
                token_hits += 1
            elif tok in t_norm:
                score += 15
                token_hits += 1
            else:
                for kw in item.get("keywords", []):
                    if tok in normalize_persian_search_term(kw):
                        score += 12
                        token_hits += 1
                        break

        # Bonus if all tokens matched
        if len(q_tokens) > 1 and token_hits == len(q_tokens):
            score += 40

        if score > 0:
            results.append((score, -idx, item))

    # Sort descending by relevance score, with tie-break preserving registry order
    results.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return [it for _, _, it in results[:limit]]
