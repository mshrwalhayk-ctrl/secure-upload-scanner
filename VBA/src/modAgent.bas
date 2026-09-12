Attribute VB_Name = "modAgent"
'====================================================================
' modAgent  -  حلقة الوكيل: يفهم الطلب، ينفّذ أدوات على الشيت،
'              يقرأ النتيجة، ويعيد الكرّة لحد ما يخلّص، ثم يرد بالعربي
'====================================================================
Option Explicit

Private mHist As Collection
Private mSink As Object           ' نموذج الواجهة لعرض الحالة (اختياري)
Private mBusy As Boolean

Public Sub SetStatusSink(ByVal o As Object)
    Set mSink = o
End Sub

Private Sub Status(ByVal s As String)
    On Error Resume Next
    If Not mSink Is Nothing Then mSink.AgentStatus s
    DoEvents
End Sub

Public Function IsBusy() As Boolean
    IsBusy = mBusy
End Function

Public Sub ResetConversation()
    Set mHist = Nothing
    UndoClear
End Sub

'====================================================================
' نقطة الدخول: يرجع رد المساعد النهائي بالعربي
'====================================================================
Public Function AgentAsk(ByVal userText As String) As String
    If mBusy Then
        AgentAsk = "في طلب شغّال حالياً، استنّى يخلص."
        Exit Function
    End If
    mBusy = True
    On Error GoTo Fail

    If mHist Is Nothing Then Set mHist = NewHistory()
    TrimHistory mHist, 40

    AddMsg mHist, "user", BuildContextBlock() & vbLf & vbLf & "طلب المستخدم:" & vbLf & userText

    Dim native As Boolean, toolsJson As String
    native = SupportsNativeTools()
    If native Then toolsJson = ToolsJson(CfgAllowMacros())

    Dim stepNo As Long, maxSteps As Long, badCount As Long
    maxSteps = CfgMaxSteps()

    Dim finalText As String

    For stepNo = 1 To maxSteps
        Status "أفكّر... (خطوة " & stepNo & " من " & maxSteps & ")"

        Dim rep As AiReply
        rep = AiChat(SystemPrompt(native), mHist, toolsJson)

        If Not rep.Ok Then
            ' لو النموذج ما بيدعم الأدوات الأصلية، منرجع للوضع النصي تلقائياً
            If native And InStr(rep.ErrText, "استدعاء الأدوات") > 0 Then
                native = False
                toolsJson = ""
                Status "النموذج ما بيدعم الأدوات الأصلية — بحوّل للوضع النصي..."
                GoTo NextStep
            End If
            mBusy = False
            AgentAsk = "⚠ " & rep.ErrText
            Exit Function
        End If

        '================= الوضع الأصلي (Function Calling) =================
        If native Then
            Dim nCalls As Long
            nCalls = 0
            If Not rep.ToolCalls Is Nothing Then nCalls = rep.ToolCalls.Count

            If nCalls = 0 Then
                finalText = Trim$(rep.Text)
                Exit For
            End If

            AddAssistantToolCalls mHist, rep.RawToolCalls, rep.Text

            Dim k As Long
            For k = 1 To nCalls
                Dim call_ As Object
                Set call_ = rep.ToolCalls.Item(k)

                Dim tName As String, tArgsRaw As String
                tName = CStr(call_.Item("name"))
                tArgsRaw = CStr(call_.Item("args"))

                Status "أنفّذ: " & ToolLabel(tName) & " ..."

                Dim argsObj As Object
                Set argsObj = ParseArgs(tArgsRaw)

                Dim result As String
                If argsObj Is Nothing Then
                    result = "خطأ: تعذّر قراءة معاملات الأداة." & vbLf & _
                             "اللي وصلني: " & Left$(tArgsRaw, 200) & vbLf & _
                             "أعد إرسال الاستدعاء بمعاملات صحيحة."
                Else
                    result = ToolExecute(tName, argsObj)
                End If

                If Len(result) > 12000 Then result = Left$(result, 12000) & vbLf & "...(مقتطع)"
                AddToolResult mHist, CStr(call_.Item("id")), result
            Next k

            GoTo NextStep
        End If

        '================= الوضع النصي الاحتياطي =================
        AddMsg mHist, "assistant", rep.Text

        Dim toolName As String, argsObj2 As Object, rawBody As String
        Dim verdict As Long
        verdict = ExtractAction(rep.Text, toolName, argsObj2, rawBody)

        Select Case verdict
            Case ACT_OK
                badCount = 0
                Status "أنفّذ: " & ToolLabel(toolName) & " ..."
                Dim res2 As String
                res2 = ToolExecute(toolName, argsObj2)
                If Len(res2) > 12000 Then res2 = Left$(res2, 12000) & vbLf & "...(مقتطع)"

                Dim followUp As String
                followUp = "نتيجة الأداة [" & toolName & "]:" & vbLf & res2
                If CountMarkers(rep.Text) > 1 Then
                    followUp = followUp & vbLf & vbLf & _
                        "تنبيه: بعتّ أكتر من كتلة أداة في رد واحد. نفّذت الأولى فقط." & vbLf & _
                        "ابعت الباقي وحدة وحدة وانتظر نتيجة كل وحدة."
                End If
                followUp = followUp & vbLf & vbLf & _
                    "الآن ابعت الخطوة التالية: كتلة أداة واحدة فقط، " & _
                    "أو ردك النهائي بالعربي إذا خلصت كل الخطوات."
                AddMsg mHist, "user", followUp

            Case ACT_BAD
                badCount = badCount + 1
                If badCount >= 3 Then
                    finalText = "⚠ النموذج مش ملتزم بصيغة استدعاء الأدوات، فما انعمل شي على الملف." & vbLf & vbLf & _
                                "آخر محاولة منه:" & vbLf & Left$(rawBody, 300) & vbLf & vbLf & _
                                "الحل الأفضل: بدّل المزوّد لـ OmniRoute أو NVIDIA من الإعدادات — " & _
                                "هدول بيدعموا استدعاء الأدوات الأصلي وما بتصير هالمشكلة نهائياً."
                    Exit For
                End If
                Status "بصحّح صيغة الاستدعاء... (محاولة " & badCount & " من 3)"
                AddMsg mHist, "user", BadFormatMessage(rawBody)

            Case Else
                finalText = Trim$(rep.Text)
                Exit For
        End Select

NextStep:
    Next stepNo

    If Len(finalText) = 0 Then
        finalText = "وصلت للحد الأقصى من الخطوات (" & maxSteps & ") ولسا ما خلّصت." & vbLf & _
                    "جرّب تقسّم الطلب لخطوات أصغر، أو زيد عدد الخطوات من الإعدادات."
    End If

    Status ""
    mBusy = False
    AgentAsk = finalText
    Exit Function

Fail:
    mBusy = False
    Status ""
    AgentAsk = "⚠ صار خطأ غير متوقع: " & Err.Description
End Function

'--- قراءة معاملات الأداة القادمة من المزوّد ------------------------
Private Function ParseArgs(ByVal raw As String) As Object
    On Error GoTo Fail
    Dim t As String
    t = Trim$(raw)
    If Len(t) = 0 Then
        Set ParseArgs = CreateObject("Scripting.Dictionary")
        Exit Function
    End If

    Dim j As Variant
    Set j = JsonTryParse(t)

    ' بعض النماذج بتغلّف المعاملات بعلامات تنصيص مرتين
    If j Is Nothing Then
        If Left$(t, 1) = """" And Right$(t, 1) = """" Then
            Dim inner As Variant
            inner = JsonTryParse(t)
            Set j = JsonTryParse(Mid$(t, 2, Len(t) - 2))
        End If
    End If

    If j Is Nothing Then Set ParseArgs = Nothing: Exit Function
    If TypeName(j) <> "Dictionary" Then Set ParseArgs = Nothing: Exit Function
    Set ParseArgs = j
    Exit Function
Fail:
    Set ParseArgs = Nothing
End Function

Private Function ToolLabel(ByVal t As String) As String
    Select Case LCase$(t)
        Case "read_range":        ToolLabel = "قراءة البيانات"
        Case "write_values":      ToolLabel = "كتابة قيم"
        Case "write_formula":     ToolLabel = "كتابة معادلة"
        Case "format_range":      ToolLabel = "تنسيق"
        Case "calc":              ToolLabel = "حساب دقيق"
        Case "precise_sum":       ToolLabel = "جمع دقيق"
        Case "list_sheets":       ToolLabel = "استعراض الأوراق"
        Case "add_sheet":         ToolLabel = "إضافة ورقة"
        Case "find_text":         ToolLabel = "بحث"
        Case "sort_range":        ToolLabel = "فرز"
        Case "remove_duplicates": ToolLabel = "حذف المكرر"
        Case "create_chart":      ToolLabel = "رسم بياني"
        Case "goal_seek":         ToolLabel = "بحث عن هدف"
        Case "stats_summary":     ToolLabel = "ملخص إحصائي"
        Case "check_column":      ToolLabel = "تدقيق عمود"
        Case "trial_balance":     ToolLabel = "ميزان المراجعة"
        Case "benford_analysis":  ToolLabel = "تحليل بنفورد"
        Case "find_gaps":         ToolLabel = "فحص التسلسل"
        Case "reconcile":         ToolLabel = "مطابقة"
        Case "aging_analysis":    ToolLabel = "أعمار الديون"
        Case "duplicate_payments": ToolLabel = "المدفوعات المكررة"
        Case "run_vba":           ToolLabel = "تشغيل ماكرو"
        Case Else:                ToolLabel = t
    End Select
End Function

'====================================================================
' استخراج أمر الأداة من رد النموذج
' الصيغة المتفق عليها:
'   <<<TOOL>>>
'   { "tool": "...", "args": { ... } }
'   <<<END>>>
' وبنقبل كمان كتلة ```json فيها نفس الشكل (تسامحاً مع النماذج الضعيفة)
'====================================================================
Public Const ACT_NONE As Long = 0      ' ما في استدعاء أداة -> هذا رد نهائي
Public Const ACT_OK   As Long = 1      ' استدعاء سليم
Public Const ACT_BAD  As Long = 2      ' حاول يستدعي أداة بصيغة غلط -> نصحّحه

Private Function ExtractAction(ByVal txt As String, ByRef toolName As String, _
                               ByRef argsObj As Object, ByRef rawBody As String) As Long
    Dim body As String
    Dim attempted As Boolean

    body = ExtractToolBody(txt)
    attempted = (Len(Trim$(body)) > 0)

    If Not attempted Then
        body = FencedJson(txt)
        attempted = (Len(Trim$(body)) > 0)
    End If

    ' حتى لو ما في علامات: إذا الرد ذكر اسم أداة مع قوس، فهي محاولة استدعاء فاشلة
    If Not attempted Then
        If LooksLikeToolAttempt(txt) Then
            rawBody = Left$(Trim$(txt), 400)
            ExtractAction = ACT_BAD
            Exit Function
        End If
        ExtractAction = ACT_NONE
        Exit Function
    End If

    rawBody = Left$(Trim$(body), 400)

    ' قص أي كلام حوالين الـ JSON: من أول { لآخر }
    Dim core As String
    core = Trim$(body)
    Dim p1 As Long, p2 As Long
    p1 = InStr(core, "{")
    p2 = InStrRev(core, "}")
    If p1 > 0 And p2 > p1 Then core = Mid$(core, p1, p2 - p1 + 1)

    Dim j As Variant
    Set j = JsonTryParse(core)
    If j Is Nothing Then ExtractAction = ACT_BAD: Exit Function
    If TypeName(j) <> "Dictionary" Then ExtractAction = ACT_BAD: Exit Function
    If Not j.Exists("tool") Then ExtractAction = ACT_BAD: Exit Function

    toolName = Trim$(CStr(j.Item("tool")))
    If Len(toolName) = 0 Then ExtractAction = ACT_BAD: Exit Function

    Set argsObj = Nothing
    If j.Exists("args") Then
        If IsObject(j.Item("args")) Then
            If TypeName(j.Item("args")) = "Dictionary" Then
                Set argsObj = j.Item("args")
            Else
                ' args إجت كقائمة قيم بدل كائن -> صيغة غلط
                ExtractAction = ACT_BAD
                Exit Function
            End If
        End If
    End If
    If argsObj Is Nothing Then Set argsObj = CreateObject("Scripting.Dictionary")

    ExtractAction = ACT_OK
End Function

'--- هل الرد فيه محاولة استدعاء أداة بدون العلامات الصحيحة؟ ---------
Private Function LooksLikeToolAttempt(ByVal txt As String) As Boolean
    Dim names As Variant, i As Long
    names = Array("read_range", "write_values", "write_formula", "format_range", _
                  "list_sheets", "add_sheet", "find_text", "sort_range", _
                  "remove_duplicates", "create_chart", "precise_sum", "run_vba", "calc", _
                  "goal_seek", "stats_summary", "check_column")
    For i = LBound(names) To UBound(names)
        Dim p As Long
        p = InStr(1, txt, names(i), vbTextCompare)
        If p > 0 Then
            ' اسم أداة متبوع مباشرة (خلال 40 حرف) بقوس معقوف
            Dim q As Long
            q = InStr(p, txt, "{")
            If q > 0 And (q - p) <= 40 Then LooksLikeToolAttempt = True: Exit Function
        End If
    Next i
    If InStr(1, txt, "<<<TOOL", vbTextCompare) > 0 Then LooksLikeToolAttempt = True
End Function

Private Function CountMarkers(ByVal txt As String) As Long
    Dim n As Long, p As Long
    p = 1
    Do
        p = InStr(p, txt, "<<<TOOL", vbTextCompare)
        If p = 0 Then Exit Do
        n = n + 1
        p = p + 7
    Loop
    CountMarkers = n
End Function

'--- رسالة التصحيح اللي بترجع للنموذج -------------------------------
Private Function BadFormatMessage(ByVal rawBody As String) As String
    Dim s As String
    s = "خطأ: صيغة استدعاء الأداة غلط، فما انعمل ولا شي على الملف." & vbLf & vbLf
    s = s & "اللي بعتّه:" & vbLf & Left$(rawBody, 300) & vbLf & vbLf
    s = s & "الصيغة الصحيحة، كتلة وحدة بس، JSON صحيح، بمفتاحين ""tool"" و ""args""،" & vbLf
    s = s & "و ""args"" كائن بأسماء الحقول مش قائمة قيم:" & vbLf & vbLf
    s = s & "<<<TOOL>>>" & vbLf
    s = s & "{""tool"":""read_range"",""args"":{""sheet"":""Sheet1"",""address"":""A1:D50""}}" & vbLf
    s = s & "<<<END>>>" & vbLf & vbLf
    s = s & "أعد إرسال أول أداة بس بالصيغة الصحيحة، بدون أي كلام قبلها أو بعدها، وانتظر نتيجتها."
    BadFormatMessage = s
End Function

'--- استخراج متسامح: بيقبل <<<TOOL>>> و <<<TOOL> و <<<TOOL>> --------
Private Function ExtractToolBody(ByVal s As String) As String
    Dim p As Long, q As Long, e As Long, ch As String
    p = InStr(1, s, "<<<TOOL", vbTextCompare)
    If p = 0 Then Exit Function

    q = p + Len("<<<TOOL")
    Do While q <= Len(s)
        ch = Mid$(s, q, 1)
        If ch = ">" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            q = q + 1
        Else
            Exit Do
        End If
    Loop

    e = InStr(q, s, "<<<END", vbTextCompare)
    If e = 0 Then e = Len(s) + 1
    ExtractToolBody = Mid$(s, q, e - q)
End Function

Private Function BetweenMarkers(ByVal s As String, ByVal m1 As String, ByVal m2 As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, s, m1, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(m1)
    p2 = InStr(p1, s, m2, vbTextCompare)
    If p2 = 0 Then p2 = Len(s) + 1
    BetweenMarkers = Mid$(s, p1, p2 - p1)
End Function

Private Function FencedJson(ByVal s As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, s, "```json", vbTextCompare)
    If p1 > 0 Then
        p1 = p1 + 7
    Else
        p1 = InStr(s, "```")
        If p1 = 0 Then Exit Function
        p1 = p1 + 3
    End If
    p2 = InStr(p1, s, "```")
    If p2 = 0 Then Exit Function
    Dim inner As String
    inner = Trim$(Mid$(s, p1, p2 - p1))
    If Left$(inner, 1) = "{" And InStr(inner, """tool""") > 0 Then FencedJson = inner
End Function

'====================================================================
' سياق الملف اللي بينشاف للنموذج مع كل طلب
'====================================================================
Public Function BuildContextBlock() As String
    On Error GoTo Simple

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then BuildContextBlock = "[لا يوجد ملف إكسل مفتوح حالياً]": Exit Function
    If wb.Name = ThisWorkbook.Name Then BuildContextBlock = "[ملف العمل غير مفتوح]": Exit Function

    Dim sb As String
    sb = "[سياق الملف الحالي]" & vbLf
    sb = sb & "الملف: " & wb.Name & vbLf
    sb = sb & "الأوراق: "

    Dim ws As Worksheet, first As Boolean
    first = True
    For Each ws In wb.Worksheets
        If Not first Then sb = sb & " ، "
        sb = sb & ws.Name
        first = False
    Next ws
    sb = sb & vbLf

    Dim aws As Object
    Set aws = wb.ActiveSheet
    If TypeName(aws) = "Worksheet" Then
        sb = sb & "الورقة النشطة: " & aws.Name
        On Error Resume Next
        sb = sb & "   النطاق المستخدم: " & aws.UsedRange.Address(False, False)
        sb = sb & "   (" & aws.UsedRange.Rows.Count & " صف × " & aws.UsedRange.Columns.Count & " عمود)"
        On Error GoTo Simple
        sb = sb & vbLf

        ' عناوين الأعمدة (أول صف من النطاق المستخدم)
        Dim hdr As String
        hdr = FirstRowsPreview(aws, 4)
        If Len(hdr) > 0 Then sb = sb & "معاينة أول صفوف:" & vbLf & hdr
    End If

    On Error Resume Next
    Dim sel As Object
    Set sel = Application.Selection
    If TypeName(sel) = "Range" Then
        sb = sb & "التحديد الحالي: " & sel.Worksheet.Name & "!" & sel.Address(False, False) & vbLf
        If sel.Cells.Count = 1 Then
            sb = sb & "  محتوى الخلية: " & Left$(CStr(sel.Text), 200)
            If Left$(sel.Formula, 1) = "=" Then sb = sb & "   المعادلة: " & sel.Formula
            sb = sb & vbLf
        End If
    End If
    On Error GoTo Simple

    BuildContextBlock = sb
    Exit Function

Simple:
    BuildContextBlock = "[تعذّر قراءة سياق الملف]"
End Function

Private Function FirstRowsPreview(ByVal ws As Object, ByVal nRows As Long) As String
    On Error GoTo Fail
    Dim ur As Range
    Set ur = ws.UsedRange
    Dim r As Long, c As Long, sb As String
    Dim maxR As Long, maxC As Long
    maxR = ur.Rows.Count: If maxR > nRows Then maxR = nRows
    maxC = ur.Columns.Count: If maxC > 15 Then maxC = 15

    For r = 1 To maxR
        sb = sb & "  " & ur.Cells(r, 1).Address(False, False) & ": "
        For c = 1 To maxC
            If c > 1 Then sb = sb & " | "
            sb = sb & Left$(CStr(ur.Cells(r, c).Text), 22)
        Next c
        If ur.Columns.Count > maxC Then sb = sb & " | ..."
        sb = sb & vbLf
        If Len(sb) > 2500 Then Exit For
    Next r
    FirstRowsPreview = sb
    Exit Function
Fail:
    FirstRowsPreview = ""
End Function

'====================================================================
' التعليمات الأساسية للنموذج
'====================================================================
Public Function SystemPrompt(Optional ByVal native As Boolean = False) As String
    Dim s As String

    s = "أنت ""مساعد المحاسب"" — مساعد ذكي مدمج داخل مايكروسوفت إكسل، بتشتغل مع محاسب مالي عربي." & vbLf & _
        "شخصيتك: خبير محاسبة ومالية ورياضيات، دقيق جداً بالأرقام، عملي، وبتحكي عربي بسيط ومباشر." & vbLf & vbLf

    s = s & "== قواعد أساسية ==" & vbLf & _
        "1. كل ردودك النهائية بالعربي. أسماء دوال إكسل ومراجع الخلايا تبقى إنجليزي كما هي." & vbLf & _
        "2. ممنوع تحسب أي عملية حسابية بعقلك. أي جمع/طرح/ضرب/قسمة/نسبة/فائدة → استخدم أداة calc." & vbLf & _
        "   حتى لو العملية بسيطة. الأرقام المالية ما بتحتمل خطأ." & vbLf & _
        "3. ما تفترض شكل البيانات. اقرأ الشيت أولاً بـ read_range أو list_sheets قبل ما تعدّل." & vbLf & _
        "4. المعادلات تُكتب بأسماء إنجليزية وبفاصلة عادية: =SUMIFS(C:C,A:A,""x"")" & vbLf & _
        "5. لما تخلّص، اشرح للمحاسب باختصار شو عملت ووين، وإذا في معادلة اشرح منطقها بسطر." & vbLf & _
        "6. إذا الطلب ناقص أو غامض بشكل بيأثر على النتيجة، اسأل سؤال واحد واضح بدل ما تخمّن." & vbLf & _
        "7. لا تعِد كتابة بيانات المستخدم كاملة في ردك — اكتفِ بالخلاصة." & vbLf & vbLf

    If native Then
        ' الأدوات معرّفة عند المزوّد نفسه، فما بنحتاج نشرح صيغة نصية.
        s = s & "== الأدوات ==" & vbLf & _
            "عندك مجموعة أدوات جاهزة بتشتغل على الشيت مباشرة. استدعِها بالطريقة الرسمية" & vbLf & _
            "(tool calling) — مش بكتابة JSON داخل النص." & vbLf & _
            "• استدعِ أداة وحدة بكل خطوة وانتظر نتيجتها قبل الخطوة اللي بعدها." & vbLf & _
            "• ممنوع تقول ""تمّ"" أو ""أضفت"" قبل ما توصلك نتيجة الأداة فعلياً." & vbLf & _
            "• نتيجة read_range ممكن تكون مقتطعة — اعتمد على عدد الصفوف المذكور" & vbLf & _
            "  في السطر الأول، مش على عدد الصفوف المعروضة، لما تحدد نطاق المعادلة." & vbLf & vbLf

        s = s & "== دليل المدقق المحترف ==" & vbLf & _
            "أنت مش بس بتكتب معادلات — أنت بتفكّر زي مدقق عنده 10 سنين خبرة." & vbLf & _
            "لما يعطيك المحاسب بيانات، اسأل حالك: وين ممكن يكون الخطأ؟ واستخدم الأداة المناسبة:" & vbLf & _
            "• ""الميزان ما بيوزن"" → trial_balance. بيقلك الفرق ويقترح سببه" & vbLf & _
            "  (فرق بيقبل القسمة على 9 = خطأ تبديل أرقام، ونص الفرق = قيد بالجهة الغلط)." & vbLf & _
            "• ""في أرقام مشبوهة"" أو تدقيق مصاريف → benford_analysis. بيكشف الأرقام" & vbLf & _
            "  المُدخلة يدوياً والتحايل على حدود الموافقة." & vbLf & _
            "• فواتير أو شيكات أو سندات → find_gaps. الأرقام الناقصة والمكررة بالتسلسل." & vbLf & _
            "• ""طابق كشف البنك مع الدفتر"" → reconcile. بيقلك وين الفروقات بالضبط." & vbLf & _
            "• ذمم مدينة → aging_analysis. وبينبّه لو فوق 90 يوم تجاوزت 25٪." & vbLf & _
            "• مصاريف أو مدفوعات موردين → duplicate_payments." & vbLf & _
            "• قبل أي مجموع → stats_summary لكشف الأرقام المخزّنة كنص." & vbLf & vbLf & _
            "عادات المدقق المحترف:" & vbLf & _
            "- ما تقبل رقم بدون ما تتحقق من مصدره." & vbLf & _
            "- لما تلاقي مشكلة، قل حجمها بالأرقام (كم صف، كم مبلغ، كم نسبة) مش بس ""في مشكلة""." & vbLf & _
            "- فرّق بين ""مؤشر يستحق الفحص"" و""خطأ مؤكد"". ما تتهم بدون دليل." & vbLf & _
            "- رتّب النتائج حسب الأثر المالي، الأكبر أولاً." & vbLf & _
            "- إذا الطلب عام مثل ""دقّق الشيت""، شغّل أكتر من أداة فحص وبعدين لخّص." & vbLf & vbLf

        s = s & "== أسلوب الشغل الصحيح ==" & vbLf & _
            "- اقرأ → احسب بـ calc → اكتب → نسّق → لخّص." & vbLf & _
            "- المعادلة تُكتب على النطاق كامل مرة وحدة (مثل E2:E420) مش خلية خلية." & vbLf & _
            "- بعد أي write_formula، اقرأ النتيجة اللي رجعتلك وتأكد إنها منطقية." & vbLf & _
            "- قبل ما تعتمد على مجموع عمود، شغّل stats_summary — إذا في أرقام مخزّنة كنص" & vbLf & _
            "  المجموع بيطلع ناقص وأنت ما بتحس. نبّه المحاسب واقترح تصليحها." & vbLf & _
            "- في أسئلة ""شو لازم يكون X عشان Y يصير كذا"" استخدم goal_seek مش التخمين." & vbLf & _
            "- في الشغل المحاسبي: انتبه للتقريب، للنصوص اللي شكلها أرقام، وللصفوف الفاضية."

        SystemPrompt = s
        Exit Function
    End If

    s = s & "== كيف تستدعي أداة (اقرأ هذا القسم بعناية) ==" & vbLf & _
        "لما تحتاج تشوف أو تعدّل الشيت، ارجع بهذا الشكل بالضبط ولا شي غيره:" & vbLf & _
        "<<<TOOL>>>" & vbLf & _
        "{""tool"":""read_range"",""args"":{""sheet"":""اليومية"",""address"":""A1:D420""}}" & vbLf & _
        "<<<END>>>" & vbLf & vbLf & _
        "قواعد صارمة ما فيها استثناء:" & vbLf & _
        "• المحتوى بين العلامتين لازم يكون JSON صحيح فيه مفتاحين بالضبط: ""tool"" و ""args""." & vbLf & _
        "• ""args"" لازم يكون كائن بأسماء الحقول، مش قائمة قيم:" & vbLf & _
        "    صح  : {""tool"":""read_range"",""args"":{""sheet"":""اليومية"",""address"":""A1:D420""}}" & vbLf & _
        "    غلط : read_range {""اليومية"", ""A1:D420""}" & vbLf & _
        "    غلط : {""tool"":""read_range"",""args"":[""اليومية"",""A1:D420""]}" & vbLf & _
        "• أداة **واحدة فقط** في الرد. ممنوع كتلتين أو أكثر." & vbLf & _
        "• ممنوع أي كلام قبل <<<TOOL>>> أو بعد <<<END>>>." & vbLf & _
        "• بعد ما ترسل الأداة، توقّف وانتظر. نتيجتها رح توصلك في الرسالة الجاية، وبعدها ترسل التالية." & vbLf & _
        "• ممنوع تقول ""تمّ"" أو ""أضفت"" أو ""حسبت"" قبل ما توصلك نتيجة الأداة فعلياً." & vbLf & _
        "  نتيجة الأداة هي الدليل الوحيد إنه الشي انعمل. بدونها ما صار شي على الملف." & vbLf & vbLf

    s = s & "مثال كامل على مهمة من عدة خطوات (لاحظ التبادل خطوة بخطوة):" & vbLf & _
        "  أنت   : <<<TOOL>>>{""tool"":""read_range"",""args"":{""sheet"":""اليومية"",""address"":""A1:D5""}}<<<END>>>" & vbLf & _
        "  النظام: نتيجة الأداة [read_range]: ... البيانات ..." & vbLf & _
        "  أنت   : <<<TOOL>>>{""tool"":""write_values"",""args"":{""sheet"":""اليومية"",""address"":""E1"",""values"":[[""ضريبة 16%""]]}}<<<END>>>" & vbLf & _
        "  النظام: نتيجة الأداة [write_values]: تم كتابة 1 صف × 1 عمود في اليومية!E1." & vbLf & _
        "  أنت   : <<<TOOL>>>{""tool"":""write_formula"",""args"":{""sheet"":""اليومية"",""address"":""E2:E420"",""formula"":""=C2*0.16""}}<<<END>>>" & vbLf & _
        "  النظام: نتيجة الأداة [write_formula]: تم وضع المعادلة ... أول نتيجة: 200" & vbLf & _
        "  أنت   : خلصت. ضفت عمود الضريبة في العمود E والمجموع في E421." & vbLf & vbLf & _
        "لما تخلص كل الخطوات فعلاً، ارجع بنص عربي عادي بدون <<<TOOL>>> — وهذا ردك النهائي." & vbLf & vbLf

    s = s & "== الأدوات المتاحة ==" & vbLf

    s = s & "read_range   {sheet, address}" & vbLf & _
        "   قراءة نطاق. مثال: {""sheet"":""Sheet1"",""address"":""A1:F30""}. اترك address فاضي = التحديد الحالي." & vbLf

    s = s & "list_sheets  {}" & vbLf & _
        "   أسماء كل الأوراق ونطاقاتها المستخدمة." & vbLf

    s = s & "calc         {expression, sheet}" & vbLf & _
        "   حساب دقيق بمحرّك إكسل. مثال: {""expression"":""SUMIFS(C2:C500,A2:A500,\""مبيعات\"")""}" & vbLf & _
        "   بيقبل أي دالة إكسل: SUM, ROUND, PMT, IRR, NPV, XNPV, VLOOKUP ... وبيقبل مراجع خلايا." & vbLf

    s = s & "precise_sum  {sheet, address}" & vbLf & _
        "   جمع بدقة 28 خانة عشرية — استخدمه لمطابقة الأرصدة والمبالغ الحساسة." & vbLf

    s = s & "write_values {sheet, address, values}" & vbLf & _
        "   كتابة قيم. values مصفوفة صفوف: [[""البيان"",""المبلغ""],[""إيجار"",1200]]" & vbLf & _
        "   address = الخلية الأولى فقط (مثل ""H1"")." & vbLf

    s = s & "write_formula {sheet, address, formula}" & vbLf & _
        "   كتابة معادلة على نطاق. مثال: {""address"":""E2:E500"",""formula"":""=C2*D2""}" & vbLf & _
        "   إكسل بيعدّل المراجع النسبية لكل صف تلقائياً." & vbLf

    s = s & "format_range {sheet, address, number_format, bold, italic, font_size, font_color," & vbLf & _
        "              fill, align, wrap, borders, autofit, freeze_header}" & vbLf & _
        "   الألوان بصيغة ""#RRGGBB"". number_format مثل ""#,##0.00"" أو ""0.0%"" أو ""dd/mm/yyyy""." & vbLf

    s = s & "add_sheet    {name}" & vbLf
    s = s & "find_text    {text, sheet}" & vbLf
    s = s & "sort_range   {sheet, address, key_column, order:asc|desc, has_header}" & vbLf
    s = s & "remove_duplicates {sheet, address, columns:""1,2"", has_header}" & vbLf
    s = s & "create_chart {sheet, data_address, chart_type:column|line|pie|bar|scatter|area, title, anchor}" & vbLf & vbLf

    s = s & "stats_summary {sheet, address}" & vbLf & _
        "   ملخص دقيق لنطاق: العدد، المجموع، المتوسط، الوسيط، الانحراف المعياري،" & vbLf & _
        "   وأهم شي: كم خلية فيها رقم مخزّن كنص (سبب رقم 1 لاختلاف المجاميع بالمحاسبة)." & vbLf

    s = s & "check_column {sheet, address}" & vbLf & _
        "   تدقيق عمود: أرقام كنص، مسافات زائدة، قيم خطأ، وعدد القيم المكررة." & vbLf

    s = s & "goal_seek {sheet, target_cell, target_value, changing_cell}" & vbLf & _
        "   البحث عن هدف: بيلاقي القيمة اللازمة في خلية عشان معادلة بخلية تانية توصل لرقم معيّن." & vbLf & _
        "   مثال: شو لازم يكون سعر البيع (B2) عشان صافي الربح (B10) يصير 25000." & vbLf & _
        "   شرط: خلية الهدف لازم تكون معادلة بتعتمد على الخلية المتغيّرة." & vbLf

    s = s & "run_vba      {code, purpose}" & vbLf & _
        "   آخر حل فقط، لما ما في أداة بتكفي (جداول محورية، معالجة معقدة، تكرار طويل)." & vbLf & _
        "   المستخدم بيشوف الكود وبيوافق قبل التشغيل، وبتنحفظ نسخة احتياطية." & vbLf & _
        "   اكتب الكود كامل داخل Sub، بدون MsgBox، وبدون تغيير ملفات تانية." & vbLf & vbLf

    s = s & "== أسلوب الشغل الصحيح ==" & vbLf & _
        "- اقرأ → احسب بـ calc → اكتب → نسّق → لخّص." & vbLf & _
        "- المعادلة تُكتب على النطاق كامل مرة وحدة (address = ""E2:E420"") مش خلية خلية." & vbLf & _
        "- بعد أي write_formula، اقرأ النتيجة اللي رجعتلك وتأكد إنها منطقية (مش #REF! ولا #VALUE!)." & vbLf & _
        "- إذا طلع خطأ، صلّح المعادلة وأعد المحاولة مرة أو مرتين، وإذا فشلت اشرح السبب للمحاسب." & vbLf & _
        "- في الشغل المحاسبي: انتبه للتقريب، للأصفار المخفية، للنصوص اللي شكلها أرقام، وللصفوف الفاضية." & vbLf & _
        "- قبل ما تعتمد على مجموع عمود، شغّل stats_summary عليه — إذا في أرقام مخزّنة كنص" & vbLf & _
        "  المجموع بيطلع ناقص وأنت ما بتحس. نبّه المحاسب واقترح تصليحها." & vbLf & _
        "- في أسئلة ""شو لازم يكون X عشان Y يصير كذا"" استخدم goal_seek مش التخمين." & vbLf

    SystemPrompt = s
End Function
