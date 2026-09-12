Attribute VB_Name = "modSchema"
'====================================================================
' modSchema  -  تعريف الأدوات بصيغة JSON Schema للـ Function Calling
'
' هذا هو الفرق الجوهري عن النسخة القديمة:
' بدل ما نطلب من النموذج يكتب JSON بإيده (وبيغلط)، منعطي المزوّد
' تعريف الأدوات، والمزوّد بيرجّعلنا الاستدعاء منظّم ومضمون الشكل.
'====================================================================
Option Explicit

'--- مساعد لبناء تعريف أداة واحدة ------------------------------------
Private Function Fn(ByVal name As String, ByVal desc As String, _
                    ByVal props As String, ByVal req As String) As String
    Fn = "{""type"":""function"",""function"":{" & _
         """name"":" & JsonStr(name) & "," & _
         """description"":" & JsonStr(desc) & "," & _
         """parameters"":{""type"":""object"",""properties"":{" & props & "}" & _
         IIf(Len(req) > 0, ",""required"":[" & req & "]", "") & _
         ",""additionalProperties"":false}}}"
End Function

Private Function S(ByVal desc As String) As String
    S = "{""type"":""string"",""description"":" & JsonStr(desc) & "}"
End Function

Private Function B(ByVal desc As String) As String
    B = "{""type"":""boolean"",""description"":" & JsonStr(desc) & "}"
End Function

Private Function N(ByVal desc As String) As String
    N = "{""type"":""number"",""description"":" & JsonStr(desc) & "}"
End Function

Private Function P(ByVal name As String, ByVal schema As String) As String
    P = JsonStr(name) & ":" & schema
End Function

'====================================================================
' مصفوفة الأدوات الكاملة
'====================================================================
Public Function ToolsJson(Optional ByVal allowMacros As Boolean = True) As String
    Dim t As String
    Dim SHEET As String
    SHEET = P("sheet", S("اسم الورقة. اتركه فارغاً للورقة النشطة."))

    ' ---------- قراءة واستكشاف ----------
    t = Fn("list_sheets", _
        "استعراض كل أوراق الملف ونطاقاتها المستخدمة. ابدأ فيها إذا ما بتعرف شكل الملف.", _
        "", "")

    t = t & "," & Fn("read_range", _
        "قراءة قيم ومعادلات نطاق من الشيت. اتركه فارغاً لقراءة التحديد الحالي.", _
        SHEET & "," & P("address", S("عنوان النطاق مثل A1:F50")), "")

    t = t & "," & Fn("find_text", _
        "البحث عن نص في الملف وإرجاع عناوين الخلايا اللي فيها.", _
        P("text", S("النص المطلوب")) & "," & SHEET, """text""")

    ' ---------- الحساب الدقيق ----------
    t = t & "," & Fn("calc", _
        "حساب أي تعبير رياضي أو دالة إكسل بمحرّك إكسل نفسه. " & _
        "استخدمها لكل عملية حسابية بدون استثناء — ممنوع تحسب بعقلك. " & _
        "بتقبل كل دوال إكسل: SUM, SUMIFS, ROUND, PMT, IRR, NPV, XNPV, VLOOKUP, MMULT ...", _
        P("expression", S("التعبير بدون علامة يساوي، مثل: SUMIFS(C2:C500,A2:A500,""مبيعات"")")) & "," & SHEET, _
        """expression""")

    t = t & "," & Fn("precise_sum", _
        "جمع نطاق بدقة 28 خانة عشرية لمطابقة الأرصدة الحساسة، مع تنبيه عن الأرقام المخزّنة كنص.", _
        SHEET & "," & P("address", S("النطاق المطلوب جمعه")), """address""")

    t = t & "," & Fn("stats_summary", _
        "ملخص إحصائي لنطاق: العدد، المجموع، المتوسط، الوسيط، الانحراف المعياري، " & _
        "وأهم شي: كم خلية فيها رقم مخزّن كنص (السبب الأول لاختلاف المجاميع بالمحاسبة).", _
        SHEET & "," & P("address", S("النطاق")), """address""")

    t = t & "," & Fn("check_column", _
        "تدقيق عمود: أرقام مخزّنة كنص، مسافات زائدة، قيم خطأ، وعدد القيم المكررة.", _
        SHEET & "," & P("address", S("نطاق العمود")), """address""")

    t = t & "," & Fn("goal_seek", _
        "البحث عن هدف: إيجاد القيمة اللازمة في خلية عشان معادلة بخلية تانية توصل لرقم معيّن. " & _
        "شرط: خلية الهدف لازم تكون معادلة بتعتمد على الخلية المتغيّرة.", _
        SHEET & "," & P("target_cell", S("خلية المعادلة، مثل B10")) & "," & _
        P("target_value", N("الرقم المطلوب الوصول له")) & "," & _
        P("changing_cell", S("الخلية اللي رح تتغيّر، مثل B2")), _
        """target_cell"",""target_value"",""changing_cell""")

    ' ---------- الكتابة ----------
    t = t & "," & Fn("write_values", _
        "كتابة قيم في الشيت. values مصفوفة صفوف، وaddress هي الخلية الأولى فقط.", _
        SHEET & "," & P("address", S("الخلية الأولى، مثل H1")) & "," & _
        """values"":{""type"":""array"",""description"":""مصفوفة صفوف، كل صف مصفوفة قيم""," & _
        """items"":{""type"":""array"",""items"":{}}}", _
        """address"",""values""")

    t = t & "," & Fn("write_formula", _
        "كتابة معادلة على نطاق كامل مرة وحدة. إكسل بيعدّل المراجع النسبية لكل صف تلقائياً. " & _
        "اكتب النطاق كامل (مثل E2:E420) مش خلية خلية.", _
        SHEET & "," & P("address", S("النطاق الكامل، مثل E2:E420")) & "," & _
        P("formula", S("المعادلة بأسماء إنجليزية، مثل =C2*0.16")), _
        """address"",""formula""")

    t = t & "," & Fn("format_range", _
        "تنسيق نطاق: صيغة الأرقام، خط عريض، ألوان، محاذاة، حدود، وضبط عرض الأعمدة.", _
        SHEET & "," & P("address", S("النطاق")) & "," & _
        P("number_format", S("صيغة الأرقام مثل #,##0.00 أو 0.0% أو dd/mm/yyyy")) & "," & _
        P("bold", B("خط عريض")) & "," & _
        P("font_size", N("حجم الخط")) & "," & _
        P("font_color", S("لون الخط بصيغة #RRGGBB")) & "," & _
        P("fill", S("لون الخلفية بصيغة #RRGGBB أو none")) & "," & _
        P("align", S("right أو left أو center")) & "," & _
        P("wrap", B("التفاف النص")) & "," & _
        P("borders", B("إضافة حدود")) & "," & _
        P("autofit", B("ضبط عرض الأعمدة تلقائياً")) & "," & _
        P("freeze_header", B("تجميد صف العناوين")), _
        """address""")

    ' ---------- تنظيم البيانات ----------
    t = t & "," & Fn("sort_range", _
        "فرز نطاق حسب عمود معيّن.", _
        SHEET & "," & P("address", S("النطاق")) & "," & _
        P("key_column", N("رقم العمود داخل النطاق، يبدأ من 1")) & "," & _
        P("order", S("asc تصاعدي أو desc تنازلي")) & "," & _
        P("has_header", B("هل أول صف عناوين")), """address""")

    t = t & "," & Fn("remove_duplicates", _
        "حذف الصفوف المكررة من نطاق.", _
        SHEET & "," & P("address", S("النطاق")) & "," & _
        P("columns", S("أرقام الأعمدة مفصولة بفواصل مثل 1,2 — اتركه فارغاً لكل الأعمدة")) & "," & _
        P("has_header", B("هل أول صف عناوين")), """address""")

    t = t & "," & Fn("add_sheet", "إنشاء ورقة جديدة في الملف.", _
        P("name", S("اسم الورقة")), """name""")

    t = t & "," & Fn("create_chart", _
        "إنشاء رسم بياني من نطاق بيانات.", _
        SHEET & "," & P("data_address", S("نطاق البيانات بما فيه العناوين")) & "," & _
        P("chart_type", S("column أو line أو pie أو bar أو scatter أو area")) & "," & _
        P("title", S("عنوان الرسم")) & "," & _
        P("anchor", S("خلية مكان الرسم، اختياري")), """data_address""")

    ' ---------- أدوات المدقق المحترف ----------
    t = t & "," & Fn("trial_balance", _
        "فحص ميزان المراجعة: هل مجموع المدين = مجموع الدائن؟ " & _
        "وإذا في فرق بيقترح سببه (خطأ تبديل أرقام، قيد بالجهة الغلط، قيد مفقود).", _
        SHEET & "," & P("debit_address", S("نطاق عمود المدين")) & "," & _
        P("credit_address", S("نطاق عمود الدائن")), _
        """debit_address"",""credit_address""")

    t = t & "," & Fn("benford_analysis", _
        "تحليل بنفورد لكشف الأرقام المفبركة أو المُدخلة يدوياً. " & _
        "بيقارن توزيع أول رقم معنوي مع التوزيع الطبيعي المتوقع. " & _
        "بدّه 50 قيمة على الأقل. تقنية تدقيق معيارية.", _
        SHEET & "," & P("address", S("نطاق المبالغ")), """address""")

    t = t & "," & Fn("find_gaps", _
        "فحص تسلسل أرقام الفواتير أو الشيكات أو السندات: " & _
        "بيرجّع الأرقام الناقصة والمكررة. الفجوة ممكن تعني مستند محذوف أو ما انسجل.", _
        SHEET & "," & P("address", S("نطاق عمود الأرقام التسلسلية")), """address""")

    t = t & "," & Fn("reconcile", _
        "مطابقة عمودين من المبالغ (مثل كشف البنك مع الدفتر) وتحديد " & _
        "القيم اللي ما لها مقابل في كل جهة، مع فرق المجموع.", _
        SHEET & "," & P("address_a", S("النطاق الأول")) & "," & _
        P("address_b", S("النطاق الثاني")) & "," & _
        P("tolerance", N("هامش المطابقة، افتراضي 0.01")), _
        """address_a"",""address_b""")

    t = t & "," & Fn("aging_analysis", _
        "تحليل أعمار الذمم: توزيع المبالغ على فئات 30/60/90/180/أكثر " & _
        "مع النسب، وتنبيه إذا الديون المتقادمة تجاوزت 25٪.", _
        SHEET & "," & P("date_address", S("نطاق عمود التواريخ")) & "," & _
        P("amount_address", S("نطاق عمود المبالغ")) & "," & _
        P("as_of", S("تاريخ الاحتساب، افتراضي اليوم")), _
        """date_address"",""amount_address""")

    t = t & "," & Fn("duplicate_payments", _
        "كشف المدفوعات المكررة: نفس الجهة بنفس المبلغ أكثر من مرة.", _
        SHEET & "," & P("key_address", S("عمود اسم المورد أو الجهة")) & "," & _
        P("amount_address", S("عمود المبالغ")), _
        """key_address"",""amount_address""")

    ' ---------- الماكرو ----------
    If allowMacros Then
        t = t & "," & Fn("run_vba", _
            "تشغيل كود VBA على الملف. آخر حل فقط لما ما في أداة بتكفي " & _
            "(جداول محورية، معالجة معقدة، تكرار طويل). " & _
            "المستخدم بيشوف الكود وبيوافق قبل التشغيل، وبتنحفظ نسخة احتياطية. " & _
            "اكتب الكود كامل داخل Sub، بدون MsgBox، وبدون فتح أو تعديل ملفات تانية.", _
            P("code", S("كود VBA كامل داخل Sub")) & "," & _
            P("purpose", S("شرح بسطر واحد بالعربي شو بيعمل الكود")), _
            """code"",""purpose""")
    End If

    ToolsJson = "[" & t & "]"
End Function
