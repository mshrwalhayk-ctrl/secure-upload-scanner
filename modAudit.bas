Attribute VB_Name = "modAudit"
'====================================================================
' modAudit  -  أدوات المدقق المالي المحترف
'
' هاي الأدوات اللي بتفرق بين "مساعد بيكتب معادلات" وبين
' "مدقق بخبرة ١٠ سنين". كلها تقنيات فعلية بتستخدم بالتدقيق المهني.
'====================================================================
Option Explicit

'--------------------------------------------------------------------
' 1) ميزان المراجعة: هل المدين = الدائن؟
'--------------------------------------------------------------------
Public Function AU_TrialBalance(ByVal args As Object) As String
    Dim ws As Object, dr As Object, cr As Object
    Set ws = AuSheet(AuArg(args, "sheet"))
    Set dr = ws.Range(AuAddr(AuArg(args, "debit_address")))
    Set cr = ws.Range(AuAddr(AuArg(args, "credit_address")))

    Dim sumD As Variant, sumC As Variant, nD As Long, nC As Long
    sumD = AuSum(dr, nD)
    sumC = AuSum(cr, nC)

    Dim diff As Variant
    diff = CDec(sumD) - CDec(sumC)

    Dim s As String
    s = "ميزان المراجعة — " & ws.Name & vbLf
    s = s & "مجموع المدين : " & NumToText(sumD) & "   (" & nD & " قيد)" & vbLf
    s = s & "مجموع الدائن : " & NumToText(sumC) & "   (" & nC & " قيد)" & vbLf
    s = s & "الفرق        : " & NumToText(diff) & vbLf & vbLf

    If Abs(CDbl(diff)) < 0.005 Then
        s = s & "✔ الميزان مطابق."
    Else
        s = s & "✖ الميزان غير مطابق بفرق " & NumToText(Abs(CDbl(diff))) & vbLf
        s = s & "الفحوصات المقترحة بالترتيب:" & vbLf
        If Abs(CDbl(diff)) Mod 9 = 0 Then
            s = s & "  • الفرق بيقبل القسمة على 9 ← غالباً خطأ تبديل أرقام (transposition)" & vbLf
        End If
        Dim half As Double
        half = Abs(CDbl(diff)) / 2
        s = s & "  • دوّر على قيد بقيمة " & NumToText(half) & " متسجّل بالجهة الغلط" & vbLf
        s = s & "  • دوّر على قيد بقيمة " & NumToText(Abs(CDbl(diff))) & " مفقود من جهة" & vbLf
        s = s & "  • تأكد ما في أرقام مخزّنة كنص (شغّل stats_summary على العمودين)"
    End If
    AU_TrialBalance = s
End Function

'--------------------------------------------------------------------
' 2) تحليل بنفورد: كشف الأرقام المفبركة
'--------------------------------------------------------------------
Public Function AU_Benford(ByVal args As Object) As String
    Dim ws As Object, rng As Object
    Set ws = AuSheet(AuArg(args, "sheet"))
    Set rng = ws.Range(AuAddr(AuArg(args, "address")))

    Dim cnt(1 To 9) As Long
    Dim total As Long
    Dim cel As Object

    For Each cel In rng.Cells
        If IsNumeric(cel.Value) And Not IsEmpty(cel.Value) Then
            Dim v As Double
            v = Abs(CDbl(cel.Value))
            If v > 0 Then
                Do While v < 1
                    v = v * 10
                Loop
                Do While v >= 10
                    v = v / 10
                Loop
                Dim d As Long
                d = Int(v)
                If d >= 1 And d <= 9 Then
                    cnt(d) = cnt(d) + 1
                    total = total + 1
                End If
            End If
        End If
        If total > 200000 Then Exit For
    Next cel

    If total < 50 Then
        AU_Benford = "تحليل بنفورد بدّه 50 رقم على الأقل (المتاح: " & total & ")." & vbLf & _
                     "التحليل على عيّنة صغيرة بيعطي إنذارات كاذبة."
        Exit Function
    End If

    Dim s As String, i As Long
    Dim chi As Double, maxDev As Double, maxDigit As Long
    s = "تحليل بنفورد — " & ws.Name & "!" & rng.Address(False, False) & vbLf
    s = s & "عدد القيم المفحوصة: " & total & vbLf & vbLf
    s = s & "رقم | الفعلي % | المتوقع % | الانحراف" & vbLf

    For i = 1 To 9
        Dim actP As Double, expP As Double, dev As Double
        actP = cnt(i) / total * 100
        expP = Log(1 + 1 / i) / Log(10) * 100
        dev = actP - expP
        chi = chi + ((actP - expP) ^ 2) / expP
        If Abs(dev) > Abs(maxDev) Then maxDev = dev: maxDigit = i
        s = s & " " & i & "  |  " & Format$(actP, "0.0") & "  |  " & Format$(expP, "0.0") & _
            "  |  " & Format$(dev, "+0.0;-0.0") & IIf(Abs(dev) > 5, "   ⚠", "") & vbLf
    Next i

    s = s & vbLf
    If chi < 15 Then
        s = s & "✔ التوزيع طبيعي ومتوافق مع قانون بنفورد. ما في مؤشر تلاعب."
    ElseIf chi < 30 Then
        s = s & "⚠ في انحراف متوسط (أعلى انحراف عند الرقم " & maxDigit & "). " & vbLf & _
            "ممكن يكون طبيعي إذا البيانات فيها حدود سعرية ثابتة أو أرقام مدوّرة."
    Else
        s = s & "⚠⚠ انحراف كبير عن التوزيع الطبيعي (أعلى انحراف عند الرقم " & maxDigit & ")." & vbLf & _
            "هذا مؤشر يستحق فحص يدوي: ممكن أرقام مُدخلة يدوياً، حدود موافقة متحايَل عليها، " & vbLf & _
            "أو قيم مكررة. مش دليل تلاعب لحاله — بس نقطة بداية للفحص."
    End If
    AU_Benford = s
End Function

'--------------------------------------------------------------------
' 3) فجوات التسلسل: فواتير أو شيكات أو سندات ناقصة
'--------------------------------------------------------------------
Public Function AU_FindGaps(ByVal args As Object) As String
    Dim ws As Object, rng As Object
    Set ws = AuSheet(AuArg(args, "sheet"))
    Set rng = ws.Range(AuAddr(AuArg(args, "address")))

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim mn As Double, mx As Double, first As Boolean, n As Long
    first = True

    Dim cel As Object
    For Each cel In rng.Cells
        Dim raw As String
        raw = Trim$(CStr(cel.Text))
        If Len(raw) > 0 Then
            Dim num As String, ch As String, k As Long
            num = ""
            For k = 1 To Len(raw)
                ch = Mid$(raw, k, 1)
                If ch >= "0" And ch <= "9" Then num = num & ch
            Next k
            If Len(num) > 0 And Len(num) <= 15 Then
                Dim v As Double
                v = Val(num)
                seen.Item(CStr(v)) = seen.Item(CStr(v)) + 1
                n = n + 1
                If first Then
                    mn = v: mx = v: first = False
                Else
                    If v < mn Then mn = v
                    If v > mx Then mx = v
                End If
            End If
        End If
    Next cel

    If n = 0 Then AU_FindGaps = "ما لقيت أرقام تسلسلية في النطاق.": Exit Function
    If (mx - mn) > 100000 Then
        AU_FindGaps = "المدى واسع جداً (" & NumToText(mn) & " إلى " & NumToText(mx) & ")." & vbLf & _
                      "غالباً هذا مش عمود تسلسلي. حدّد العمود الصح."
        Exit Function
    End If

    Dim gaps As String, gapCount As Long, dupCount As Long, dups As String
    Dim i As Double
    For i = mn To mx
        If Not seen.Exists(CStr(i)) Then
            gapCount = gapCount + 1
            If gapCount <= 60 Then gaps = gaps & NumToText(i) & " ، "
        ElseIf seen.Item(CStr(i)) > 1 Then
            dupCount = dupCount + 1
            If dupCount <= 30 Then dups = dups & NumToText(i) & "×" & seen.Item(CStr(i)) & " ، "
        End If
    Next i

    Dim s As String
    s = "فحص التسلسل — " & ws.Name & "!" & rng.Address(False, False) & vbLf
    s = s & "المدى: من " & NumToText(mn) & " إلى " & NumToText(mx) & _
        "   الموجود: " & n & "   المتوقع: " & NumToText(mx - mn + 1) & vbLf & vbLf

    If gapCount = 0 Then
        s = s & "✔ ما في أرقام ناقصة — التسلسل كامل." & vbLf
    Else
        s = s & "✖ ناقص " & gapCount & " رقم:" & vbLf & "   " & gaps & vbLf
        If gapCount > 60 Then s = s & "   ... (معروض أول 60)" & vbLf
        s = s & "   الفجوات بالتسلسل ممكن تعني: مستندات ملغاة، أو محذوفة، أو ما انسجلت." & vbLf
    End If

    If dupCount > 0 Then
        s = s & vbLf & "⚠ في " & dupCount & " رقم مكرر:" & vbLf & "   " & dups & vbLf & _
            "   الترقيم المكرر خطر — ممكن فاتورة مسجّلة مرتين."
    End If
    AU_FindGaps = s
End Function

'--------------------------------------------------------------------
' 4) مطابقة عمودين: وين الفروقات بالضبط
'--------------------------------------------------------------------
Public Function AU_Reconcile(ByVal args As Object) As String
    Dim ws As Object
    Set ws = AuSheet(AuArg(args, "sheet"))

    Dim ra As Object, rb As Object
    Set ra = ws.Range(AuAddr(AuArg(args, "address_a")))
    Set rb = ws.Range(AuAddr(AuArg(args, "address_b")))

    Dim tol As Double
    tol = Val(Replace(AuArg(args, "tolerance", "0.01"), ",", "."))
    If tol <= 0 Then tol = 0.01

    ' نبني قاموس من الجهة الثانية بالقيم المقرّبة
    Dim mapB As Object
    Set mapB = CreateObject("Scripting.Dictionary")
    Dim cel As Object, key As String

    For Each cel In rb.Cells
        If IsNumeric(cel.Value) And Not IsEmpty(cel.Value) Then
            key = Format$(Round(CDbl(cel.Value) / tol, 0) * tol, "0.############")
            If mapB.Exists(key) Then
                mapB.Item(key) = mapB.Item(key) & "|" & cel.Address(False, False)
            Else
                mapB.Item(key) = cel.Address(False, False)
            End If
        End If
    Next cel

    Dim unmatchedA As String, nA As Long, matched As Long
    Dim sumA As Variant, sumB As Variant
    sumA = CDec(0): sumB = CDec(0)

    For Each cel In ra.Cells
        If IsNumeric(cel.Value) And Not IsEmpty(cel.Value) Then
            nA = nA + 1
            sumA = CDec(sumA) + CDec(cel.Value)
            key = Format$(Round(CDbl(cel.Value) / tol, 0) * tol, "0.############")
            If mapB.Exists(key) Then
                matched = matched + 1
                Dim lst As String, p As Long
                lst = mapB.Item(key)
                p = InStr(lst, "|")
                If p > 0 Then mapB.Item(key) = Mid$(lst, p + 1) Else mapB.Remove key
            Else
                If nA - matched <= 50 Then
                    unmatchedA = unmatchedA & "   " & cel.Address(False, False) & " = " & _
                                 NumToText(cel.Value) & vbLf
                End If
            End If
        End If
    Next cel

    Dim unmatchedB As String, nB As Long, leftB As Long
    For Each cel In rb.Cells
        If IsNumeric(cel.Value) And Not IsEmpty(cel.Value) Then
            nB = nB + 1
            sumB = CDec(sumB) + CDec(cel.Value)
        End If
    Next cel

    Dim keys As Variant, i As Long
    keys = mapB.Keys
    For i = LBound(keys) To UBound(keys)
        Dim parts() As String, jj As Long
        parts = Split(CStr(mapB.Item(keys(i))), "|")
        For jj = 0 To UBound(parts)
            leftB = leftB + 1
            If leftB <= 50 Then unmatchedB = unmatchedB & "   " & parts(jj) & " = " & keys(i) & vbLf
        Next jj
    Next i

    Dim s As String
    s = "مطابقة " & ra.Address(False, False) & " مع " & rb.Address(False, False) & vbLf
    s = s & "الجهة أ: " & nA & " قيمة، مجموعها " & NumToText(sumA) & vbLf
    s = s & "الجهة ب: " & nB & " قيمة، مجموعها " & NumToText(sumB) & vbLf
    s = s & "فرق المجموع: " & NumToText(CDec(sumA) - CDec(sumB)) & vbLf
    s = s & "تطابق: " & matched & " قيمة   (بهامش " & NumToText(tol) & ")" & vbLf & vbLf

    If Len(unmatchedA) = 0 And Len(unmatchedB) = 0 Then
        s = s & "✔ كل القيم متطابقة."
    Else
        If Len(unmatchedA) > 0 Then
            s = s & "✖ موجود في (أ) وما له مقابل في (ب) — عددها " & (nA - matched) & ":" & vbLf & unmatchedA
        End If
        If Len(unmatchedB) > 0 Then
            s = s & "✖ موجود في (ب) وما له مقابل في (أ) — عددها " & leftB & ":" & vbLf & unmatchedB
        End If
        s = s & vbLf & "ملاحظة: قيمتين غير متطابقتين مجموعهما يساوي فرق المجموع = غالباً نفس البند بمبلغ مختلف."
    End If
    AU_Reconcile = s
End Function

'--------------------------------------------------------------------
' 5) أعمار الديون
'--------------------------------------------------------------------
Public Function AU_Aging(ByVal args As Object) As String
    Dim ws As Object
    Set ws = AuSheet(AuArg(args, "sheet"))

    Dim rd As Object, ra As Object
    Set rd = ws.Range(AuAddr(AuArg(args, "date_address")))
    Set ra = ws.Range(AuAddr(AuArg(args, "amount_address")))

    Dim asOf As Date
    Dim asStr As String
    asStr = AuArg(args, "as_of")
    If Len(asStr) > 0 And IsDate(asStr) Then asOf = CDate(asStr) Else asOf = Date

    Dim b(0 To 4) As Variant, c(0 To 4) As Long
    Dim i As Long
    For i = 0 To 4
        b(i) = CDec(0)
    Next i

    Dim rows_ As Long
    rows_ = rd.Cells.Count
    If ra.Cells.Count < rows_ Then rows_ = ra.Cells.Count

    Dim notDated As Long, total As Variant
    total = CDec(0)

    For i = 1 To rows_
        Dim dv As Variant, av As Variant
        dv = rd.Cells(i).Value
        av = ra.Cells(i).Value
        If IsNumeric(av) And Not IsEmpty(av) Then
            If IsDate(dv) Then
                Dim days As Long, idx As Long
                days = CLng(asOf - CDate(dv))
                If days <= 30 Then
                    idx = 0
                ElseIf days <= 60 Then
                    idx = 1
                ElseIf days <= 90 Then
                    idx = 2
                ElseIf days <= 180 Then
                    idx = 3
                Else
                    idx = 4
                End If
                b(idx) = CDec(b(idx)) + CDec(av)
                c(idx) = c(idx) + 1
                total = CDec(total) + CDec(av)
            Else
                notDated = notDated + 1
            End If
        End If
    Next i

    Dim nm As Variant
    nm = Array("أقل من 30 يوم", "31 - 60", "61 - 90", "91 - 180", "أكثر من 180")

    Dim s As String
    s = "تحليل أعمار الديون — كما في " & Format$(asOf, "dd/mm/yyyy") & vbLf
    s = s & "الفئة            | العدد | المبلغ | النسبة" & vbLf
    For i = 0 To 4
        Dim pct As Double
        If CDbl(total) <> 0 Then pct = CDbl(b(i)) / CDbl(total) * 100
        s = s & nm(i) & " | " & c(i) & " | " & NumToText(b(i)) & " | " & Format$(pct, "0.0") & "%" & vbLf
    Next i
    s = s & "الإجمالي: " & NumToText(total) & vbLf

    If notDated > 0 Then s = s & vbLf & "⚠ " & notDated & " صف مبلغه موجود بس تاريخه ناقص أو غير صالح." & vbLf

    Dim overdue As Double
    overdue = CDbl(b(3)) + CDbl(b(4))
    If CDbl(total) <> 0 Then
        If overdue / CDbl(total) > 0.25 Then
            s = s & vbLf & "⚠ " & Format$(overdue / CDbl(total) * 100, "0") & _
                "٪ من الذمم أعمارها فوق 90 يوم — نسبة مرتفعة تستدعي مخصص ديون مشكوك فيها."
        End If
    End If
    AU_Aging = s
End Function

'--------------------------------------------------------------------
' 6) المدفوعات المكررة: نفس المبلغ لنفس الجهة خلال فترة قصيرة
'--------------------------------------------------------------------
Public Function AU_DuplicatePayments(ByVal args As Object) As String
    Dim ws As Object
    Set ws = AuSheet(AuArg(args, "sheet"))

    Dim rk As Object, ram As Object
    Set rk = ws.Range(AuAddr(AuArg(args, "key_address")))
    Set ram = ws.Range(AuAddr(AuArg(args, "amount_address")))

    Dim rows_ As Long
    rows_ = rk.Cells.Count
    If ram.Cells.Count < rows_ Then rows_ = ram.Cells.Count
    If rows_ > 20000 Then rows_ = 20000

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim s As String, hits As Long, amt As Variant
    amt = CDec(0)

    Dim i As Long
    For i = 1 To rows_
        Dim kv As String, av As Variant
        kv = Trim$(CStr(rk.Cells(i).Text))
        av = ram.Cells(i).Value
        If Len(kv) > 0 And IsNumeric(av) And Not IsEmpty(av) Then
            Dim key As String
            key = LCase$(kv) & "#" & Format$(CDbl(av), "0.00")
            If seen.Exists(key) Then
                hits = hits + 1
                amt = CDec(amt) + CDec(av)
                If hits <= 40 Then
                    s = s & "   صف " & rk.Cells(i).Row & " يطابق صف " & seen.Item(key) & _
                        "  —  " & kv & "  بمبلغ " & NumToText(av) & vbLf
                End If
            Else
                seen.Item(key) = rk.Cells(i).Row
            End If
        End If
    Next i

    Dim out As String
    out = "فحص المدفوعات المكررة — " & ws.Name & vbLf
    out = out & "الصفوف المفحوصة: " & rows_ & vbLf & vbLf
    If hits = 0 Then
        out = out & "✔ ما في مدفوعات مكررة (نفس الجهة + نفس المبلغ)."
    Else
        out = out & "⚠ لقيت " & hits & " حالة تكرار محتملة، مجموعها " & NumToText(amt) & ":" & vbLf & s
        If hits > 40 Then out = out & "   ... (معروض أول 40)" & vbLf
        out = out & vbLf & "هاي مؤشرات مش أحكام — في حالات تكرار مشروع (دفعات شهرية ثابتة). " & _
              "راجع رقم المستند والتاريخ لكل حالة."
    End If
    AU_DuplicatePayments = out
End Function

'====================================================================
' مساعدات داخلية
'====================================================================
Private Function AuArg(ByVal args As Object, ByVal k As String, _
                       Optional ByVal dflt As String = "") As String
    On Error Resume Next
    If args Is Nothing Then AuArg = dflt: Exit Function
    If Not args.Exists(k) Then AuArg = dflt: Exit Function
    If IsObject(args.Item(k)) Then AuArg = dflt: Exit Function
    If IsNull(args.Item(k)) Then AuArg = dflt: Exit Function
    AuArg = CStr(args.Item(k))
End Function

Private Function AuSheet(ByVal nm As String) As Object
    Dim wb As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 900, , "ما في ملف إكسل مفتوح."
    If Len(Trim$(nm)) = 0 Then Set AuSheet = wb.ActiveSheet: Exit Function

    Dim ws As Object
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then Set AuSheet = ws: Exit Function
    Next ws
    Err.Raise vbObjectError + 901, , "ما في ورقة اسمها """ & nm & """"
End Function

Private Function AuAddr(ByVal a As String) As String
    a = Trim$(a)
    If InStr(a, "!") > 0 Then a = Mid$(a, InStr(a, "!") + 1)
    AuAddr = Replace(a, "$", "")
End Function

Private Function AuSum(ByVal rng As Object, ByRef n As Long) As Variant
    Dim t As Variant, cel As Object
    t = CDec(0)
    n = 0
    For Each cel In rng.Cells
        If IsNumeric(cel.Value) And Not IsEmpty(cel.Value) Then
            t = CDec(t) + CDec(cel.Value)
            n = n + 1
        End If
    Next cel
    AuSum = t
End Function
