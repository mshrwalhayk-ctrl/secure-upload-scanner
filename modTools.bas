Attribute VB_Name = "modTools"
'====================================================================
' modTools  -  الأدوات اللي بيقدر النموذج يشغّلها على ملف إكسل
'              كل أداة بترجع نص وصفي بيرجع للنموذج كنتيجة
'====================================================================
Option Explicit

Private Const MAX_READ_ROWS As Long = 300
Private Const MAX_READ_COLS As Long = 40

' مكدّس التراجع: كل عنصر = Dictionary فيه (sheet, address, formulas, desc)
Private mUndoStack As Collection

'====================================================================
' موزّع الأدوات
'====================================================================
Public Function ToolExecute(ByVal toolName As String, ByVal args As Object) As String
    On Error GoTo Fail

    Select Case LCase$(Trim$(toolName))
        Case "read_range":        ToolExecute = T_ReadRange(args)
        Case "write_values":      ToolExecute = T_WriteValues(args)
        Case "write_formula":     ToolExecute = T_WriteFormula(args)
        Case "format_range":      ToolExecute = T_FormatRange(args)
        Case "calc":              ToolExecute = T_Calc(args)
        Case "precise_sum":       ToolExecute = T_PreciseSum(args)
        Case "list_sheets":       ToolExecute = T_ListSheets(args)
        Case "add_sheet":         ToolExecute = T_AddSheet(args)
        Case "find_text":         ToolExecute = T_FindText(args)
        Case "sort_range":        ToolExecute = T_SortRange(args)
        Case "remove_duplicates": ToolExecute = T_RemoveDuplicates(args)
        Case "create_chart":      ToolExecute = T_CreateChart(args)
        Case "goal_seek":         ToolExecute = T_GoalSeek(args)
        Case "stats_summary":     ToolExecute = T_Stats(args)
        Case "check_column":      ToolExecute = T_CheckColumn(args)
        Case "run_vba":           ToolExecute = T_RunVba(args)
        Case "trial_balance":     ToolExecute = AU_TrialBalance(args)
        Case "benford_analysis":  ToolExecute = AU_Benford(args)
        Case "find_gaps":         ToolExecute = AU_FindGaps(args)
        Case "reconcile":         ToolExecute = AU_Reconcile(args)
        Case "aging_analysis":    ToolExecute = AU_Aging(args)
        Case "duplicate_payments": ToolExecute = AU_DuplicatePayments(args)
        Case Else
            ToolExecute = "خطأ: أداة غير معروفة (" & toolName & ")."
    End Select
    Exit Function

Fail:
    ToolExecute = "خطأ أثناء تنفيذ الأداة " & toolName & ": " & Err.Description
End Function

'====================================================================
' مساعدات
'====================================================================
Private Function A(ByVal args As Object, ByVal k As String, _
                   Optional ByVal dflt As String = "") As String
    On Error Resume Next
    If args Is Nothing Then A = dflt: Exit Function
    If Not args.Exists(k) Then A = dflt: Exit Function
    If IsObject(args.Item(k)) Then A = dflt: Exit Function
    If IsNull(args.Item(k)) Then A = dflt: Exit Function
    A = CStr(args.Item(k))
End Function

Private Function ABool(ByVal args As Object, ByVal k As String, _
                       Optional ByVal dflt As Boolean = False) As Boolean
    Dim s As String
    s = LCase$(A(args, k, IIf(dflt, "true", "false")))
    ABool = (s = "true" Or s = "1" Or s = "yes" Or s = "نعم")
End Function

Private Function TargetBook() As Workbook
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 800, , "ما في ملف إكسل مفتوح."
    If wb.Name = ThisWorkbook.Name Then Err.Raise vbObjectError + 801, , "افتح ملف العمل أولاً."
    Set TargetBook = wb
End Function

Private Function GetSheet(ByVal nm As String) As Worksheet
    Dim wb As Workbook
    Set wb = TargetBook()
    If Len(Trim$(nm)) = 0 Then
        If TypeName(wb.ActiveSheet) <> "Worksheet" Then
            Err.Raise vbObjectError + 802, , "الورقة النشطة ليست ورقة بيانات."
        End If
        Set GetSheet = wb.ActiveSheet
        Exit Function
    End If
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then Set GetSheet = ws: Exit Function
    Next ws
    Err.Raise vbObjectError + 803, , "ما في ورقة اسمها """ & nm & """. الأوراق الموجودة: " & SheetNames()
End Function

Private Function SheetNames() As String
    Dim ws As Worksheet, s As String
    On Error Resume Next
    For Each ws In TargetBook().Worksheets
        s = s & IIf(Len(s) > 0, " ، ", "") & ws.Name
    Next ws
    SheetNames = s
End Function

Private Function GetRange(ByVal ws As Worksheet, ByVal addr As String) As Range
    addr = Trim$(addr)
    If Len(addr) = 0 Then
        Set GetRange = ws.Application.Selection
        Exit Function
    End If
    ' إزالة اسم الورقة لو النموذج كتبه داخل العنوان
    If InStr(addr, "!") > 0 Then addr = Mid$(addr, InStr(addr, "!") + 1)
    addr = Replace(addr, "$", "")
    On Error GoTo Bad
    Set GetRange = ws.Range(addr)
    Exit Function
Bad:
    Err.Raise vbObjectError + 804, , "عنوان نطاق غير صالح: """ & addr & """"
End Function

'====================================================================
' مكدّس التراجع
'====================================================================
Private Sub PushUndo(ByVal rng As Range, ByVal desc As String)
    On Error Resume Next
    If mUndoStack Is Nothing Then Set mUndoStack = New Collection

    If rng.Cells.Count > 200000 Then Exit Sub   ' نطاق ضخم: نتخطى الحفظ

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Item("book") = rng.Worksheet.Parent.Name
    d.Item("sheet") = rng.Worksheet.Name
    d.Item("address") = rng.Address
    d.Item("desc") = desc
    d.Item("formulas") = rng.Formula
    d.Item("numfmt") = rng.NumberFormat
    mUndoStack.Add d

    Do While mUndoStack.Count > 20
        mUndoStack.Remove 1
    Loop
End Sub

Public Function UndoCount() As Long
    If mUndoStack Is Nothing Then UndoCount = 0 Else UndoCount = mUndoStack.Count
End Function

Public Function UndoLast() As String
    On Error GoTo Fail
    If mUndoStack Is Nothing Then UndoLast = "ما في شي للتراجع عنه.": Exit Function
    If mUndoStack.Count = 0 Then UndoLast = "ما في شي للتراجع عنه.": Exit Function

    Dim d As Object
    Set d = mUndoStack.Item(mUndoStack.Count)

    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.Workbooks(CStr(d.Item("book")))
    Set ws = wb.Worksheets(CStr(d.Item("sheet")))

    Application.ScreenUpdating = False
    ws.Range(CStr(d.Item("address"))).Formula = d.Item("formulas")
    On Error Resume Next
    ws.Range(CStr(d.Item("address"))).NumberFormat = d.Item("numfmt")
    On Error GoTo Fail
    Application.ScreenUpdating = True

    mUndoStack.Remove mUndoStack.Count
    UndoLast = "تم التراجع عن: " & d.Item("desc") & _
               "  (باقي " & mUndoStack.Count & " خطوة قابلة للتراجع)"
    Exit Function
Fail:
    Application.ScreenUpdating = True
    UndoLast = "تعذّر التراجع: " & Err.Description
End Function

Public Sub UndoClear()
    Set mUndoStack = Nothing
End Sub

'====================================================================
' الأدوات
'====================================================================

'--- قراءة نطاق ------------------------------------------------------
Private Function T_ReadRange(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "address"))

    Dim r1 As Long, c1 As Long, nR As Long, nC As Long
    r1 = rng.Row: c1 = rng.Column
    nR = rng.Rows.Count: nC = rng.Columns.Count

    Dim truncated As Boolean
    If nR > MAX_READ_ROWS Then nR = MAX_READ_ROWS: truncated = True
    If nC > MAX_READ_COLS Then nC = MAX_READ_COLS: truncated = True

    Dim sb As String, r As Long, c As Long, cel As Range
    sb = "الورقة: " & ws.Name & "   النطاق: " & rng.Address(False, False) & vbLf
    sb = sb & "الصيغة: صف | عمود | قيمة | معادلة (إن وجدت)" & vbLf

    For r = 0 To nR - 1
        Dim line As String
        line = ""
        For c = 0 To nC - 1
            Set cel = ws.Cells(r1 + r, c1 + c)
            Dim val As String
            If IsError(cel.Value) Then
                val = "#خطأ"
            ElseIf IsEmpty(cel.Value) Then
                val = ""
            Else
                val = CStr(cel.Text)
            End If
            If Left$(cel.Formula, 1) = "=" Then
                val = val & " {" & cel.Formula & "}"
            End If
            If c > 0 Then line = line & vbTab
            line = line & val
        Next c
        sb = sb & ws.Cells(r1 + r, c1).Address(False, False) & vbTab & line & vbLf
        If Len(sb) > 24000 Then truncated = True: Exit For
    Next r

    If truncated Then sb = sb & vbLf & "(تم اقتطاع الناتج — النطاق أكبر من المعروض)"
    T_ReadRange = sb
End Function

'--- كتابة قيم -------------------------------------------------------
Private Function T_WriteValues(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetSheet(A(args, "sheet"))

    Dim vals As Variant
    Set vals = Nothing
    On Error Resume Next
    Set vals = args.Item("values")
    On Error GoTo 0
    If vals Is Nothing Then T_WriteValues = "خطأ: الحقل values مفقود أو ليس مصفوفة.": Exit Function
    If TypeName(vals) <> "Collection" Then T_WriteValues = "خطأ: values لازم تكون مصفوفة صفوف.": Exit Function
    If vals.Count = 0 Then T_WriteValues = "خطأ: values فارغة.": Exit Function

    ' تحديد الأبعاد
    Dim nR As Long, nC As Long, i As Long, j As Long
    nR = vals.Count
    nC = 1
    For i = 1 To nR
        If TypeName(vals.Item(i)) = "Collection" Then
            If vals.Item(i).Count > nC Then nC = vals.Item(i).Count
        End If
    Next i

    Dim arr() As Variant
    ReDim arr(1 To nR, 1 To nC)
    For i = 1 To nR
        If TypeName(vals.Item(i)) = "Collection" Then
            For j = 1 To vals.Item(i).Count
                Dim cv As Variant
                cv = vals.Item(i).Item(j)
                If IsNull(cv) Then cv = ""
                arr(i, j) = cv
            Next j
        Else
            arr(i, 1) = vals.Item(i)
        End If
    Next i

    Dim startAddr As String
    startAddr = A(args, "address")
    If Len(startAddr) = 0 Then T_WriteValues = "خطأ: الحقل address مطلوب.": Exit Function

    Dim topLeft As Range, target As Range
    Set topLeft = GetRange(ws, startAddr).Cells(1, 1)
    Set target = topLeft.Resize(nR, nC)

    PushUndo target, "كتابة قيم في " & ws.Name & "!" & target.Address(False, False)

    Application.ScreenUpdating = False
    target.Value = arr
    Application.ScreenUpdating = True

    T_WriteValues = "تم كتابة " & nR & " صف × " & nC & " عمود في " & _
                    ws.Name & "!" & target.Address(False, False) & "."
End Function

'--- كتابة معادلة ----------------------------------------------------
Private Function T_WriteFormula(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))

    Dim addr As String, f As String
    addr = A(args, "address")
    f = A(args, "formula")
    If Len(addr) = 0 Then T_WriteFormula = "خطأ: الحقل address مطلوب.": Exit Function
    If Len(f) = 0 Then T_WriteFormula = "خطأ: الحقل formula مطلوب.": Exit Function
    If Left$(f, 1) <> "=" Then f = "=" & f

    Set rng = GetRange(ws, addr)
    PushUndo rng, "كتابة معادلة في " & ws.Name & "!" & rng.Address(False, False)

    On Error GoTo BadFormula
    Application.ScreenUpdating = False
    rng.Formula = f
    Application.ScreenUpdating = True
    On Error GoTo 0

    ' إرجاع النتيجة المحسوبة للتحقق
    Dim shown As String
    On Error Resume Next
    If IsError(rng.Cells(1, 1).Value) Then
        shown = "الخلية رجّعت خطأ: " & CStr(rng.Cells(1, 1).Text)
    Else
        shown = "أول نتيجة: " & CStr(rng.Cells(1, 1).Text)
    End If
    On Error GoTo 0

    T_WriteFormula = "تم وضع المعادلة " & f & " في " & ws.Name & "!" & _
                     rng.Address(False, False) & ". " & shown
    Exit Function

BadFormula:
    Application.ScreenUpdating = True
    T_WriteFormula = "المعادلة مرفوضة من إكسل: " & Err.Description & _
                     "  — تأكد من أسماء الدوال بالإنجليزي ومن الفواصل."
End Function

'--- تنسيق نطاق -----------------------------------------------------
Private Function T_FormatRange(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "address"))

    PushUndo rng, "تنسيق " & ws.Name & "!" & rng.Address(False, False)

    Application.ScreenUpdating = False
    On Error Resume Next

    Dim nf As String
    nf = A(args, "number_format")
    If Len(nf) > 0 Then rng.NumberFormat = nf

    If args.Exists("bold") Then rng.Font.Bold = ABool(args, "bold")
    If args.Exists("italic") Then rng.Font.Italic = ABool(args, "italic")

    Dim fs As String
    fs = A(args, "font_size")
    If Len(fs) > 0 Then rng.Font.Size = Val(fs)

    Dim fc As String
    fc = A(args, "font_color")
    If Len(fc) > 0 Then rng.Font.Color = HexToColor(fc)

    Dim fill As String
    fill = A(args, "fill")
    If Len(fill) > 0 Then
        If LCase$(fill) = "none" Then
            rng.Interior.Pattern = xlNone
        Else
            rng.Interior.Color = HexToColor(fill)
        End If
    End If

    Dim al As String
    al = LCase$(A(args, "align"))
    Select Case al
        Case "right":  rng.HorizontalAlignment = xlRight
        Case "left":   rng.HorizontalAlignment = xlLeft
        Case "center": rng.HorizontalAlignment = xlCenter
    End Select

    If args.Exists("wrap") Then rng.WrapText = ABool(args, "wrap")

    If ABool(args, "borders") Then
        rng.Borders.LineStyle = xlContinuous
        rng.Borders.Weight = xlThin
        rng.Borders.Color = RGB(170, 170, 170)
    End If

    If ABool(args, "autofit") Then rng.EntireColumn.AutoFit

    If ABool(args, "freeze_header") Then
        ws.Activate
        ws.Cells(rng.Row + 1, 1).Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    End If

    On Error GoTo 0
    Application.ScreenUpdating = True

    T_FormatRange = "تم تنسيق " & ws.Name & "!" & rng.Address(False, False) & "."
End Function

Private Function HexToColor(ByVal h As String) As Long
    h = Replace(Trim$(h), "#", "")
    If Len(h) <> 6 Then HexToColor = RGB(0, 0, 0): Exit Function
    On Error GoTo Bad
    HexToColor = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
    Exit Function
Bad:
    HexToColor = RGB(0, 0, 0)
End Function

'--- حساب -----------------------------------------------------------
Private Function T_Calc(ByVal args As Object) As String
    Dim expr As String, okFlag As Boolean
    expr = A(args, "expression")
    If Len(expr) = 0 Then T_Calc = "خطأ: الحقل expression مطلوب.": Exit Function

    Dim ws As Object
    Set ws = Nothing
    On Error Resume Next
    If Len(A(args, "sheet")) > 0 Then Set ws = GetSheet(A(args, "sheet"))
    If ws Is Nothing Then Set ws = ActiveWorkbook.ActiveSheet
    On Error GoTo 0

    Dim res As String
    res = MathEval(expr, okFlag, ws)
    If okFlag Then
        T_Calc = "النتيجة الدقيقة لـ (" & expr & ") = " & res
    Else
        T_Calc = res
    End If
End Function

'--- جمع عالي الدقة -------------------------------------------------
Private Function T_PreciseSum(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "address"))
    T_PreciseSum = "المجموع الدقيق لـ " & ws.Name & "!" & rng.Address(False, False) & _
                   " = " & PreciseSum(rng)
End Function

'--- قائمة الأوراق ---------------------------------------------------
Private Function T_ListSheets(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = TargetBook()
    Dim ws As Worksheet, sb As String
    sb = "الملف: " & wb.Name & vbLf
    For Each ws In wb.Worksheets
        Dim ur As String
        On Error Resume Next
        ur = ws.UsedRange.Address(False, False)
        On Error GoTo 0
        sb = sb & "- " & ws.Name & "  (النطاق المستخدم: " & ur & ")" & vbLf
    Next ws
    T_ListSheets = sb
End Function

'--- إضافة ورقة ------------------------------------------------------
Private Function T_AddSheet(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = TargetBook()
    Dim nm As String
    nm = A(args, "name", "ورقة جديدة")

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        T_AddSheet = "الورقة """ & nm & """ موجودة أصلاً، رح نستخدمها."
        Exit Function
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    On Error Resume Next
    ws.Name = nm
    On Error GoTo 0
    T_AddSheet = "تم إنشاء ورقة جديدة اسمها """ & ws.Name & """."
End Function

'--- بحث -------------------------------------------------------------
Private Function T_FindText(ByVal args As Object) As String
    Dim txt As String
    txt = A(args, "text")
    If Len(txt) = 0 Then T_FindText = "خطأ: الحقل text مطلوب.": Exit Function

    Dim wb As Workbook
    Set wb = TargetBook()

    Dim shList As Collection
    Set shList = New Collection
    If Len(A(args, "sheet")) > 0 Then
        shList.Add GetSheet(A(args, "sheet"))
    Else
        Dim w As Worksheet
        For Each w In wb.Worksheets
            shList.Add w
        Next w
    End If

    Dim sb As String, hits As Long
    Dim i As Long
    For i = 1 To shList.Count
        Dim ws As Worksheet
        Set ws = shList.Item(i)
        Dim c As Range, firstAddr As String
        Set c = ws.UsedRange.Find(What:=txt, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not c Is Nothing Then
            firstAddr = c.Address
            Do
                hits = hits + 1
                sb = sb & ws.Name & "!" & c.Address(False, False) & "  =  " & Left$(c.Text, 80) & vbLf
                Set c = ws.UsedRange.FindNext(c)
                If c Is Nothing Then Exit Do
                If hits >= 60 Then Exit Do
            Loop While c.Address <> firstAddr
        End If
        If hits >= 60 Then Exit For
    Next i

    If hits = 0 Then
        T_FindText = "ما لقيت """ & txt & """."
    Else
        T_FindText = "عدد النتائج: " & hits & vbLf & sb
    End If
End Function

'--- فرز -------------------------------------------------------------
Private Function T_SortRange(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "address"))

    Dim keyCol As String, ord As String, hasHdr As Boolean
    keyCol = A(args, "key_column", "1")
    ord = LCase$(A(args, "order", "asc"))
    hasHdr = ABool(args, "has_header", True)

    PushUndo rng, "فرز " & ws.Name & "!" & rng.Address(False, False)

    Dim kIdx As Long
    If IsNumeric(keyCol) Then
        kIdx = CLng(keyCol)
    Else
        kIdx = ws.Range(keyCol & "1").Column - rng.Column + 1
    End If
    If kIdx < 1 Then kIdx = 1
    If kIdx > rng.Columns.Count Then kIdx = 1

    Application.ScreenUpdating = False
    rng.Sort key1:=rng.Columns(kIdx), _
             Order1:=IIf(ord = "desc", xlDescending, xlAscending), _
             Header:=IIf(hasHdr, xlYes, xlNo)
    Application.ScreenUpdating = True

    T_SortRange = "تم فرز " & ws.Name & "!" & rng.Address(False, False) & _
                  " حسب العمود رقم " & kIdx & " (" & IIf(ord = "desc", "تنازلي", "تصاعدي") & ")."
End Function

'--- حذف المكرر ------------------------------------------------------
Private Function T_RemoveDuplicates(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "address"))

    PushUndo rng, "حذف مكرر من " & ws.Name & "!" & rng.Address(False, False)

    Dim before As Long
    before = Application.WorksheetFunction.CountA(rng.Columns(1))

    Dim colsSpec As String
    colsSpec = A(args, "columns")

    Application.ScreenUpdating = False
    If Len(colsSpec) = 0 Then
        Dim allCols() As Variant, i As Long
        ReDim allCols(0 To rng.Columns.Count - 1)
        For i = 0 To rng.Columns.Count - 1
            allCols(i) = i + 1
        Next i
        rng.RemoveDuplicates Columns:=(allCols), Header:=IIf(ABool(args, "has_header", True), xlYes, xlNo)
    Else
        Dim parts() As String
        parts = Split(colsSpec, ",")
        Dim cc() As Variant
        ReDim cc(0 To UBound(parts))
        For i = 0 To UBound(parts)
            cc(i) = CLng(Val(Trim$(parts(i))))
        Next i
        rng.RemoveDuplicates Columns:=(cc), Header:=IIf(ABool(args, "has_header", True), xlYes, xlNo)
    End If
    Application.ScreenUpdating = True

    Dim afterN As Long
    afterN = Application.WorksheetFunction.CountA(rng.Columns(1))
    T_RemoveDuplicates = "تم حذف المكرر. عدد الصفوف قبل: " & before & " ، بعد: " & afterN & "."
End Function

'--- رسم بياني -------------------------------------------------------
Private Function T_CreateChart(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "data_address"))

    Dim ct As Long
    Select Case LCase$(A(args, "chart_type", "column"))
        Case "line":    ct = xlLine
        Case "pie":     ct = xlPie
        Case "bar":     ct = xlBarClustered
        Case "scatter": ct = xlXYScatterLines
        Case "area":    ct = xlArea
        Case Else:      ct = xlColumnClustered
    End Select

    Dim anchor As String
    anchor = A(args, "anchor")

    Dim co As ChartObject
    Dim L As Double, T As Double
    If Len(anchor) > 0 Then
        Dim ar As Range
        Set ar = GetRange(ws, anchor)
        L = ar.Left: T = ar.Top
    Else
        L = rng.Left + rng.Width + 20
        T = rng.Top
    End If

    Set co = ws.ChartObjects.Add(Left:=L, Top:=T, Width:=440, Height:=260)
    co.Chart.SetSourceData Source:=rng
    co.Chart.ChartType = ct

    Dim ttl As String
    ttl = A(args, "title")
    If Len(ttl) > 0 Then
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = ttl
    End If

    T_CreateChart = "تم إنشاء رسم بياني في الورقة " & ws.Name & " من البيانات " & _
                    rng.Address(False, False) & "."
End Function

'--- البحث عن الهدف (Goal Seek) --------------------------------------
'    مثال: شو لازم يكون سعر البيع عشان الربح يطلع 10000؟
Private Function T_GoalSeek(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetSheet(A(args, "sheet"))

    Dim tgt As Range, chg As Range
    Set tgt = GetRange(ws, A(args, "target_cell")).Cells(1, 1)
    Set chg = GetRange(ws, A(args, "changing_cell")).Cells(1, 1)

    Dim goalVal As Double
    goalVal = Val(Replace(A(args, "target_value", "0"), ",", "."))

    If Left$(tgt.Formula, 1) <> "=" Then
        T_GoalSeek = "خطأ: خلية الهدف (" & tgt.Address(False, False) & ") لازم تكون فيها معادلة، مش رقم ثابت."
        Exit Function
    End If

    PushUndo chg, "بحث عن هدف — تغيير " & ws.Name & "!" & chg.Address(False, False)

    Dim before As String
    before = CStr(chg.Value)

    Dim okFlag As Boolean
    On Error Resume Next
    okFlag = tgt.GoalSeek(Goal:=goalVal, ChangingCell:=chg)
    On Error GoTo 0

    If Not okFlag Then
        T_GoalSeek = "ما قدر يلاقي حل. تأكد إنه خلية الهدف بتعتمد فعلاً على الخلية المتغيّرة."
        Exit Function
    End If

    T_GoalSeek = "لقيت الحل: عشان " & tgt.Address(False, False) & " تصير " & goalVal & _
                 " لازم " & chg.Address(False, False) & " تكون " & NumToText(chg.Value) & _
                 "  (كانت " & before & "). القيمة الحالية للهدف: " & CStr(tgt.Text)
End Function

'--- ملخص إحصائي دقيق لنطاق -----------------------------------------
Private Function T_Stats(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "address"))

    Dim cel As Range
    Dim n As Long, nBlank As Long, nText As Long, nErr As Long, nTextNum As Long
    Dim total As Variant, mn As Double, mx As Double, first As Boolean
    total = CDec(0)
    first = True

    For Each cel In rng.Cells
        If IsError(cel.Value) Then
            nErr = nErr + 1
        ElseIf IsEmpty(cel.Value) Then
            nBlank = nBlank + 1
        ElseIf VarType(cel.Value) = vbString Then
            nText = nText + 1
            If IsNumeric(cel.Value) Then nTextNum = nTextNum + 1
        ElseIf IsNumeric(cel.Value) Then
            n = n + 1
            total = CDec(total) + CDec(cel.Value)
            Dim d As Double
            d = CDbl(cel.Value)
            If first Then
                mn = d: mx = d: first = False
            Else
                If d < mn Then mn = d
                If d > mx Then mx = d
            End If
        End If
        If n + nBlank + nText + nErr > 200000 Then Exit For
    Next cel

    Dim sb As String
    sb = "ملخص " & ws.Name & "!" & rng.Address(False, False) & vbLf
    sb = sb & "خلايا رقمية: " & n & vbLf
    sb = sb & "المجموع الدقيق: " & NumToText(total) & vbLf
    If n > 0 Then
        sb = sb & "المتوسط: " & NumToText(CDec(total) / CDec(n)) & vbLf
        sb = sb & "أصغر قيمة: " & NumToText(mn) & "    أكبر قيمة: " & NumToText(mx) & vbLf
        On Error Resume Next
        sb = sb & "الوسيط: " & NumToText(Application.WorksheetFunction.Median(rng)) & vbLf
        sb = sb & "الانحراف المعياري: " & NumToText(Application.WorksheetFunction.StDev(rng)) & vbLf
        On Error GoTo 0
    End If
    sb = sb & "خلايا فاضية: " & nBlank & vbLf
    sb = sb & "خلايا نصية: " & nText
    If nTextNum > 0 Then
        sb = sb & "   ⚠ منها " & nTextNum & " رقم مخزّن كنص (ما بيدخل بالمجاميع!)"
    End If
    sb = sb & vbLf & "خلايا فيها خطأ: " & nErr
    T_Stats = sb
End Function

'--- فحص تدقيقي لعمود ------------------------------------------------
Private Function T_CheckColumn(ByVal args As Object) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheet(A(args, "sheet"))
    Set rng = GetRange(ws, A(args, "address"))

    Dim cel As Range, sb As String, issues As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    For Each cel In rng.Cells
        Dim note As String
        note = ""

        If IsError(cel.Value) Then
            note = "قيمة خطأ: " & CStr(cel.Text)
        ElseIf VarType(cel.Value) = vbString Then
            If IsNumeric(cel.Value) Then
                note = "رقم مخزّن كنص: """ & cel.Text & """"
            ElseIf InStr(cel.Value, "  ") > 0 Or cel.Value <> Trim$(cel.Value) Then
                note = "مسافات زائدة في النص"
            End If
        End If

        If Len(note) > 0 Then
            issues = issues + 1
            If issues <= 40 Then
                sb = sb & "  " & cel.Address(False, False) & " — " & note & vbLf
            End If
        End If

        ' كشف التكرار
        If Not IsEmpty(cel.Value) And Not IsError(cel.Value) Then
            Dim k As String
            k = CStr(cel.Value)
            If seen.Exists(k) Then
                seen.Item(k) = seen.Item(k) + 1
            Else
                seen.Item(k) = 1
            End If
        End If

        If cel.Row - rng.Row > 100000 Then Exit For
    Next cel

    Dim dups As Long, i As Long
    Dim ks As Variant
    ks = seen.Keys
    For i = LBound(ks) To UBound(ks)
        If seen.Item(ks(i)) > 1 Then dups = dups + 1
    Next i

    Dim head As String
    head = "تدقيق " & ws.Name & "!" & rng.Address(False, False) & vbLf
    head = head & "عدد المشاكل: " & issues & "    قيم مكررة: " & dups & vbLf
    If issues > 40 Then head = head & "(معروض أول 40 مشكلة)" & vbLf
    If issues = 0 And dups = 0 Then head = head & "ما في مشاكل ظاهرة." & vbLf

    T_CheckColumn = head & sb
End Function

'====================================================================
' تشغيل كود VBA مولّد — بموافقة المستخدم فقط
'====================================================================
Private Function T_RunVba(ByVal args As Object) As String
    If Not CfgAllowMacros() Then
        T_RunVba = "تشغيل الماكرو معطّل من إعدادات الإضافة. المستخدم لازم يفعّله أولاً."
        Exit Function
    End If

    Dim code As String, purpose As String
    code = A(args, "code")
    purpose = A(args, "purpose", "(بدون وصف)")
    If Len(code) = 0 Then T_RunVba = "خطأ: الحقل code مطلوب.": Exit Function

    ' --- موافقة المستخدم ---
    Dim preview As String
    preview = code
    If Len(preview) > 2500 Then preview = Left$(preview, 2500) & vbCrLf & "... (مقتطع)"

    Dim ans As VbMsgBoxResult
    ans = MsgBox("المساعد بدّه يشغّل ماكرو على الملف." & vbCrLf & vbCrLf & _
                 "الهدف: " & purpose & vbCrLf & _
                 String(52, "-") & vbCrLf & preview & vbCrLf & String(52, "-") & vbCrLf & vbCrLf & _
                 "ملاحظة: التراجع (Ctrl+Z) ما بشتغل بعد الماكرو، بس رح نحفظ نسخة احتياطية." & vbCrLf & vbCrLf & _
                 "بدك نشغّله؟", _
                 vbYesNo + vbQuestion + vbMsgBoxRight, APP_TITLE & " — تأكيد تشغيل ماكرو")

    If ans <> vbYes Then
        T_RunVba = "المستخدم رفض تشغيل الماكرو. جرّب طريقة تانية بدون run_vba."
        Exit Function
    End If

    ' --- نسخة احتياطية ---
    Dim bkMsg As String
    bkMsg = BackupActiveBook()

    ' --- الحقن والتشغيل ---
    Dim wb As Workbook
    Set wb = TargetBook()

    Dim vbc As Object
    Dim procName As String
    procName = "MaliAI_Temp_" & Format$(Now, "hhnnss") & Int(Rnd() * 900 + 100)

    On Error GoTo NoVbom
    Set vbc = wb.VBProject.VBComponents.Add(1)      ' 1 = وحدة قياسية
    On Error GoTo Cleanup

    Dim wrapped As String
    If InStr(1, code, "Sub ", vbTextCompare) > 0 Or InStr(1, code, "Function ", vbTextCompare) > 0 Then
        ' الكود فيه إجراءات جاهزة: نلفّه بإجراء يستدعي أول إجراء
        wrapped = code
        Dim firstProc As String
        firstProc = FirstProcName(code)
        If Len(firstProc) = 0 Then
            On Error Resume Next
            wb.VBProject.VBComponents.Remove vbc
            T_RunVba = "ما قدرت أحدد اسم الإجراء داخل الكود. اكتبه داخل Sub واضح."
            Exit Function
        End If
        procName = firstProc
    Else
        wrapped = "Public Sub " & procName & "()" & vbCrLf & code & vbCrLf & "End Sub"
    End If

    vbc.CodeModule.AddFromString wrapped

    Application.Run "'" & wb.Name & "'!" & procName

    Dim compName As String
    compName = vbc.Name
    wb.VBProject.VBComponents.Remove vbc

    T_RunVba = "تم تنفيذ الماكرو بنجاح. " & bkMsg
    Exit Function

NoVbom:
    T_RunVba = "تعذّر تشغيل الماكرو: صلاحية الوصول لمشروع VBA غير مفعّلة." & vbCrLf & _
               "فعّلها من: File ← Options ← Trust Center ← Trust Center Settings ← " & _
               "Macro Settings ← ✔ Trust access to the VBA project object model."
    Exit Function

Cleanup:
    Dim eDesc As String
    eDesc = Err.Description
    On Error Resume Next
    If Not vbc Is Nothing Then wb.VBProject.VBComponents.Remove vbc
    T_RunVba = "الماكرو وقع بخطأ: " & eDesc & vbCrLf & bkMsg & vbCrLf & _
               "حلّل الخطأ وجرّب كود مصحّح أو طريقة تانية."
End Function

Private Function FirstProcName(ByVal code As String) As String
    Dim arrLines() As String, i As Long, ln As String, p As Long
    arrLines = Split(Replace(code, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(arrLines)
        ln = Trim$(arrLines(i))
        If LCase$(Left$(ln, 11)) = "public sub " Then
            FirstProcName = ProcNameFrom(Mid$(ln, 12)): Exit Function
        ElseIf LCase$(Left$(ln, 4)) = "sub " Then
            FirstProcName = ProcNameFrom(Mid$(ln, 5)): Exit Function
        ElseIf LCase$(Left$(ln, 12)) = "private sub " Then
            FirstProcName = ProcNameFrom(Mid$(ln, 13)): Exit Function
        End If
    Next i
End Function

Private Function ProcNameFrom(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, "(")
    If p > 0 Then s = Left$(s, p - 1)
    ProcNameFrom = Trim$(s)
End Function

'--- نسخة احتياطية ---------------------------------------------------
Public Function BackupActiveBook() As String
    On Error GoTo Fail
    If Not CfgAutoBackup() Then BackupActiveBook = "": Exit Function

    Dim wb As Workbook
    Set wb = TargetBook()

    Dim folder As String, base As String, ext As String, dest As String
    If Len(wb.Path) > 0 Then
        folder = wb.Path
    Else
        folder = Environ$("TEMP")
    End If

    base = wb.Name
    Dim dotPos As Long
    dotPos = InStrRev(base, ".")
    If dotPos > 0 Then
        ext = Mid$(base, dotPos)
        base = Left$(base, dotPos - 1)
    Else
        ext = ".xlsx"
    End If

    dest = folder & "\" & base & "_نسخة_احتياطية_" & Format$(Now, "yyyymmdd_hhnnss") & ext
    wb.SaveCopyAs dest
    BackupActiveBook = "نسخة احتياطية محفوظة: " & dest
    Exit Function
Fail:
    BackupActiveBook = "(تعذّر حفظ نسخة احتياطية: " & Err.Description & ")"
End Function
