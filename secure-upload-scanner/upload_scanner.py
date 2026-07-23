# -*- coding: utf-8 -*-
"""
filecheck.py — طبقة أمان: فحص سلامة الملفات المرفوعة قبل قبولها.
يرفض حسب قواعد عامة (اسم + امتداد + محتوى + حجم) لا حسب توقيع محدد،
فيمسك التهديدات الجديدة كما القديمة. لا يُنفّذ أي ملف إطلاقاً — فحص فقط.
"""
import os
import re

MAX_SIZE = 25 * 1024 * 1024  # 25 ميغابايت

# قائمة بيضاء: الامتدادات المسموحة فقط — أي شيء آخر مرفوض
ALLOWED_EXT = {
    ".xlsx", ".xlsm", ".xls", ".csv",
    ".pdf", ".docx", ".doc", ".txt", ".md",
    ".png", ".jpg", ".jpeg", ".webp", ".gif",
}

# امتدادات تنفيذية/خطرة صريحة — مرفوضة ولو ظهرت في اسم مزدوج (فاتورة.pdf.exe)
DANGEROUS_EXT = {
    ".exe", ".bat", ".cmd", ".com", ".scr", ".pif", ".msi", ".dll", ".msc",
    ".js", ".jse", ".vbs", ".vbe", ".ps1", ".sh", ".jar", ".apk", ".hta",
    ".wsf", ".reg", ".lnk", ".php", ".py", ".rb", ".pl", ".jsp", ".asp",
}

# علامة فيروس الاختبار EICAR (مبنية من أجزاء كي لا يُعلَّم هذا الملف نفسه كخبيث)
_EICAR_MARK = b"EICAR-STANDARD-" + b"ANTIVIRUS-TEST-FILE"

# بصمات بداية ملف تنفيذي (متنكّر كمستند)
_EXE_MAGIC = (b"MZ", b"\x7fELF")


def safe_name(filename):
    """يرجّع اسم الملف بلا أي مسار — يحيّد path traversal."""
    return os.path.basename(str(filename or "").replace("\\", "/").strip())


def scan_disk_file(path):
    """يفحص ملفاً موجوداً على القرص (لا مرفوعاً) بنفس منطق فحص الرفع.
    يُستدعى قبل أن يقرأ/يفتح الوكيل أي ملف، فيغطّي الملفات المنسوخة يدوياً
    للمجلدات أو الواصلة عبر الواتس — أي ملف لم يمرّ عبر بوابة /api/upload."""
    try:
        with open(path, "rb") as f:
            data = f.read(MAX_SIZE + 1)
    except Exception as e:
        return False, f"تعذّر قراءة الملف للفحص الأمني: {e}"
    return validate_upload(os.path.basename(path), data)


def validate_upload(filename, data):
    """يفحص ملفاً مرفوعاً. يرجّع (ok: bool, reason: str). data = بايتات المحتوى."""
    raw = str(filename or "")

    # 1) اسم يحوي مساراً = محاولة path traversal
    if "/" in raw or "\\" in raw or ".." in raw:
        return False, "اسم ملف يحتوي مساراً غير مسموح (path traversal)."
    name = safe_name(raw)
    if not name:
        return False, "اسم ملف غير صالح."
    low = name.lower()

    # 2) امتداد خطر في أي مقطع (يكشف الامتداد المزدوج)
    for seg in low.split(".")[1:]:
        if ("." + seg) in DANGEROUS_EXT:
            return False, f"نوع ملف خطر (.{seg}) — مرفوض."

    # 3) الامتداد النهائي ضمن القائمة البيضاء فقط
    ext = os.path.splitext(low)[1]
    if ext not in ALLOWED_EXT:
        return False, f"امتداد غير مسموح ({ext or 'بدون امتداد'}) — المسموح: مستندات وجداول وصور فقط."

    # 4) الحجم
    n = len(data or b"")
    if n == 0:
        return False, "ملف فارغ."
    if n > MAX_SIZE:
        return False, f"الملف كبير جداً ({n // (1024 * 1024)}MB) — الحد {MAX_SIZE // (1024 * 1024)}MB."

    # 5) المحتوى: فيروس اختبار EICAR (في أي مكان بالملف) + بصمة تنفيذية متنكّرة
    if _EICAR_MARK in data:
        return False, "اكتُشف توقيع فيروس الاختبار (EICAR) — مرفوض."
    head = data[:8]
    for magic in _EXE_MAGIC:
        if head.startswith(magic):
            return False, "محتوى الملف تنفيذي متنكّر كمستند — مرفوض."

    # 6) فحص عميق حسب النوع: أوفيس الحديث (أرشيف ZIP) + PDF + مستندات أوفيس القديمة
    if ext in (".xlsx", ".xlsm", ".docx"):
        ok, why = _check_office_zip(data)
        if not ok:
            return False, why
    if ext == ".pdf":
        dehexed = _pdf_dehex(data)                       # يفكّ تمويه أسماء PDF بالهكس (/J#61vaScript → /JavaScript)
        for mark in _PDF_ACTIVE:
            if mark in data or mark in dehexed:
                return False, ("ملف PDF يحوي محتوى نشطاً محقوناً "
                               f"({mark.decode()}: جافاسكربت/تشغيل/تضمين) — مرفوض.")
    if ext in (".doc", ".xls") and (b"V\x00B\x00A" in data or b"M\x00a\x00c\x00r\x00o\x00s" in data):
        return False, "مستند أوفيس قديم يحوي ماكرو VBA — مرفوض."

    if ext == ".csv":                                     # 🛡️ حقن صيغ CSV (DDE/أوامر تُنفَّذ عند الفتح بإكسل)
        ok, why = _scan_csv(data)
        if not ok:
            return False, why

    if ext in (".png", ".jpg", ".jpeg", ".webp", ".gif"):     # 🛡️ قنبلة بكسل/EXIF محقونة
        ok, why = scan_image(data)
        if not ok:
            return False, why

    return True, "الملف آمن."


# أسماء أفعال PDF النشطة (تُفحص كنص خام كما تفعل أدوات pdfid) — لا مكان لها في مستند HR
_PDF_ACTIVE = (b"/JavaScript", b"/JS", b"/Launch", b"/EmbeddedFile", b"/RichMedia", b"/XFA")

# فكّ تمويه أسماء PDF بالهكس: PDF يسمح بترميز أي حرف في الاسم كـ #XX (فـ /J#61vaScript = /JavaScript)
_PDF_HEX = re.compile(rb"#([0-9A-Fa-f]{2})")


def _pdf_dehex(data):
    """يفكّ ترميز #XX في بايتات PDF كي يُكشف الاسم المموّه (يُفحص إضافةً للنص الأصلي)."""
    try:
        return _PDF_HEX.sub(lambda m: bytes([int(m.group(1), 16)]), data)
    except Exception:
        return data


# أنماط حقن الصيغ (CSV / DDE داخل أوفيس): تُنفّذ أمراً أو تسرّب بيانات عند الفتح بإكسل
_FORMULA_INJECT = (b"cmd|", b"msexcel|", b"powershell", b"|'/c", b'|"/c',
                   b"=hyperlink(", b"=webservice(", b"=importxml(", b"=importdata(", b"=dde(", b"ddeauto")

# نفس الأنماط لفحص أجزاء إكسل الحديثة (xlsx): XML يخزّن الصيغة **بلا** علامة «=» البادئة
# (<f>WEBSERVICE(...)</f>)، فنطابق باسم الدالة وحده. يشمل التسريب عبر الجلب الخارجي.
_OFFICE_FORMULA_INJECT = (b"cmd|", b"msexcel|", b"powershell", b"|'/c", b'|"/c', b"dde(", b"ddeauto",
                          b"webservice(", b"importxml(", b"importdata(", b"importhtml(",
                          b"importrange(", b"importfeed(", b"hyperlink(", b"rtd(")


def _scan_csv(data):
    """يفحص CSV بحثاً عن حقن صيغ (=cmd|، DDE، HYPERLINK/WEBSERVICE للتسريب) — يُرفض ما يُنفّذ عند الفتح."""
    low = bytes(data[:512 * 1024]).lower()
    for m in _FORMULA_INJECT:
        if m in low:
            return False, ("ملف CSV يحوي صيغة حقن قد تُنفّذ أمراً أو تسرّب بيانات عند فتحه بإكسل "
                           f"({m.decode(errors='ignore')}) — مرفوض.")
    return True, ""


# بصمة أرشيف حقيقي: ملفات أوفيس الحديثة يجب أن تكون ZIP صالحاً يبدأ بـ PK
_ZIP_MAGIC = b"PK"


def _check_office_zip(data):
    """يفتح ملف الأوفيس كأرشيف ZIP (بلا تنفيذ) ويفحص كل جزء داخله:
    ماكرو VBA، كائنات مضمّنة (OLE)، ملف تنفيذي/امتداد خطر مدسوس، EICAR، وقنبلة ضغط."""
    import io
    import zipfile
    if not data.startswith(_ZIP_MAGIC):
        return False, "ملف أوفيس مزيف — ليس أرشيف ZIP صالحاً."
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except Exception:
        return False, "ملف أوفيس تالف — تعذّر فتح الأرشيف."
    total = 0
    for info in zf.infolist():
        n = info.filename.replace("\\", "/").lower()
        total += info.file_size
        if total > 200 * 1024 * 1024:
            return False, "أرشيف منتفخ عند فك الضغط (zip bomb) — مرفوض."
        if "vbaproject" in n:
            return False, "المستند يحوي ماكرو VBA — مرفوض."
        if "/embeddings/" in n or n.startswith("embeddings/"):
            return False, "المستند يحوي كائناً مضمّناً (OLE embedding) — مرفوض."
        if "externallink" in n:                          # رابط بيانات خارجي (قناة تسريب/DDE)
            return False, "المستند يحوي رابط بيانات خارجي (external link) — مرفوض."
        # 🛡️ تهريب الخطوط: خط مخصّص مدمج (TTF/OTF) لا مكان له في ملف HR — يُستخدم لتمويه الحروف
        if n.endswith((".ttf", ".otf", ".fon", ".eot")) or "/fonts/" in n:
            return False, "المستند يحوي خطاً مخصّصاً مدمجاً (font smuggling) — مرفوض."
        for seg in n.rsplit("/", 1)[-1].split(".")[1:]:
            if ("." + seg) in DANGEROUS_EXT:
                return False, f"المستند يحوي بداخله ملفاً خطراً (.{seg}) — مرفوض."
        try:
            part = zf.open(info).read(2 * 1024 * 1024)   # أول 2MB من كل جزء تكفي للبصمات
        except Exception:
            continue
        if _EICAR_MARK in part:
            return False, "اكتُشف توقيع فيروس الاختبار (EICAR) داخل المستند — مرفوض."
        for magic in _EXE_MAGIC:
            if part.startswith(magic):
                return False, "المستند يحوي ملفاً تنفيذياً مدسوساً بداخله — مرفوض."
        # 🛡️ استغلال المُفسِّر XML (XXE / قنبلة الكيانات): كيان خارجي أو DOCTYPE ممنوع في أوفيس سليم
        if n.endswith(".xml") or n.endswith(".rels"):
            low = part.lstrip()[:4096].lower()
            if b"<!doctype" in low or b"<!entity" in low or b"system \"" in low or b"system '" in low:
                return False, "المستند يحوي تعريف كيان XML خارجي (XXE / قنبلة كيانات) — مرفوض."
        pl = part.lower()                                # 🛡️ صيغة DDE/أمر أو تسريب بيانات داخل ورقة الإكسل
        for mrk in _OFFICE_FORMULA_INJECT:
            if mrk in pl:
                return False, ("المستند يحوي صيغة قد تُنفّذ أمراً أو تسرّب بيانات عند الفتح بإكسل "
                               f"({mrk.decode(errors='ignore')}) — مرفوض.")
    return True, ""


def _image_meta_text(im):
    """يستخرج نصّ الميتاداتا **الفعليّ** لصورة PIL (EXIF + وسوم PNG النصّية + تعليق JPEG) — حيث يعيش نصّ المستخدم،
    لا بكسلاتها. نفحص هذا (لا البايتات الخام المضغوطة) فنكشف الحقن الحقيقيّ **بلا إيجابيات كاذبة**."""
    parts = []
    try:
        ex = im.getexif()                        # EXIF المُفكَّك (قيمٌ نصّية: الطراز/البرنامج/التعليق/الوصف…)
        parts.extend(ex.values())
        for ifd in (0x8769, 0x8825):             # Exif IFD + GPS IFD (تعليقات المستخدم قد تسكن هنا)
            try:
                parts.extend(ex.get_ifd(ifd).values())
            except Exception:
                pass
    except Exception:
        pass
    try:                                         # وسوم PNG النصّية (tEXt/iTXt) وتعليق JPEG — تُخزَّن نصوصاً في info
        for k, v in (getattr(im, "info", {}) or {}).items():
            if k in ("icc_profile", "exif", "photoshop", "adobe", "dpi", "chromaticity"):
                continue                         # كتلٌ ثنائية غير نصّية → تُتجاهَل (لا معنى لفحصها)
            parts.append(v)
    except Exception:
        pass
    out = []
    for p in parts:
        try:
            out.append(p.decode("utf-8", "ignore") if isinstance(p, (bytes, bytearray)) else str(p))
        except Exception:
            continue
    return " ".join(out).lower()


def scan_image(data):
    """يفحص صورة مرفوعة قبل معالجتها بالرؤية: قنبلة فك ضغط + أبعاد مهولة + **حقن في الميتاداتا** + حمولة polyglot.
    يرجّع (ok, reason). حدود بنيوية لا توقيع. **الحقن يُفحَص في الميتاداتا الفعلية (EXIF/الوسوم) لا بكسلات الصورة**
    (فحص البكسلات المضغوطة يرفض ~٩٠٪ من الصور الحقيقية زوراً — علامةٌ قصيرة كـ`$(` تظهر بالصدفة في أي صورة)."""
    import io
    try:
        from PIL import Image
    except Exception:
        return True, ""                          # لا PIL = لا فحص صور (لا نكسر الرفع)
    Image.MAX_IMAGE_PIXELS = 200_000_000         # سقف سخيّ (يمرّر صور موبايل 108MP، يرفض القنابل)
    # ① بنية سليمة + أبعاد: **verify مباشرةً بعد open** (قيد PIL: أيّ عمليةٍ قبله — كـgetexif — تكسره).
    try:
        with Image.open(io.BytesIO(data)) as im:
            w, h = im.size
            if w * h > 150_000_000 or w > 30000 or h > 30000:   # قنبلة حقيقية = مليارات البكسل
                return False, f"صورة بأبعاد مهولة ({w}x{h}) — قنبلة فك ضغط محتملة، مرفوضة."
            im.verify()                          # يكشف الصور المشوّهة/الخبيثة بنيوياً
    except Exception as e:
        return False, f"صورة تالفة أو خطرة (قد تكون قنبلة بكسل): {e}"
    # ② الميتاداتا: **إعادة فتح** (verify أبطل الكائن الأوّل — لا يُستخرَج منه بعده)
    meta = ""
    try:
        with Image.open(io.BytesIO(data)) as im2:
            meta = _image_meta_text(im2)
    except Exception:
        meta = ""                                # تعذّرت القراءة (نادر بعد نجاح ①) → لا ميتاداتا لفحصها
    # 1) حقن في **الميتاداتا الفعلية** (نصّ EXIF/الوسوم — حيث يُحقَن المستخدم): مجموعة أوسع (نصٌّ قصير مضبوط، لا بكسلات)
    for mark in ("<?php", "<script", "javascript:", "union select", "; rm ", "eval(", "system(", "$(", "`"):
        if mark in meta:
            return False, "بيانات وصف الصورة (EXIF) تحوي نمط حقن — مرفوضة."
    # 2) حمولة polyglot ملحقة بالبايتات الخام — علامات **طويلة فقط** (≥5 بايت، آمنة من التصادم العشوائيّ مع البكسل)
    lo = data.lower()
    for mark in (b"<?php", b"<script", b"<%eval", b"passthru(", b"shell_exec"):
        if mark in lo:
            return False, "الملف يحوي حمولةً تنفيذية ملحقة (polyglot) — مرفوض."
    return True, ""
