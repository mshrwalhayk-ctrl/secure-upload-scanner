Attribute VB_Name = "modMath"
'====================================================================
' modMath  -  محرّك الحساب الدقيق
'
' الفكرة: النماذج اللغوية بتغلط بالحساب الذهني. فبدل ما نثق فيها،
' كل عملية حسابية بتنفّذها الإضافة نفسها عبر محرّك إكسل الأصلي
' (نفس المحرّك اللي بيحسب المعادلات بالشيت) + نوع Decimal بدقة
' 28 خانة للأرقام المالية. هيك النتيجة بتطلع مطابقة تماماً لإكسل.
'====================================================================
Option Explicit

'--------------------------------------------------------------------
' تقييم تعبير حسابي أو معادلة إكسل كاملة
' أمثلة:  "1250*1.16" , "SUM(A1:A20)" , "PMT(0.05/12,60,-25000)"
'         "ROUND(AVERAGE(Sheet1!B2:B99),2)"
'--------------------------------------------------------------------
Public Function MathEval(ByVal expr As String, ByRef okFlag As Boolean, _
                         Optional ByVal ctxSheet As Object = Nothing) As String
    Dim s As String
    Dim v As Variant

    okFlag = False
    s = NormalizeExpr(expr)
    If Len(s) = 0 Then
        MathEval = "تعبير فارغ."
        Exit Function
    End If

    On Error GoTo Fail

    Dim prevCalc As Long
    If ctxSheet Is Nothing Then
        v = Application.Evaluate(s)
    Else
        v = ctxSheet.Evaluate(s)
    End If

    If IsError(v) Then
        MathEval = "خطأ في التعبير: " & ErrName(v)
        Exit Function
    End If

    If IsArray(v) Then
        MathEval = ArrayToText(v)
        okFlag = True
        Exit Function
    End If

    If IsNumeric(v) Then
        MathEval = NumToText(v)
    Else
        MathEval = CStr(v)
    End If
    okFlag = True
    Exit Function

Fail:
    MathEval = "تعذّر حساب التعبير (" & Err.Description & ")."
End Function

'--------------------------------------------------------------------
' تنظيف التعبير قبل التقييم
'--------------------------------------------------------------------
Private Function NormalizeExpr(ByVal s As String) As String
    Dim i As Long, c As String, code As Long, sb As String

    s = Trim$(s)
    If Len(s) = 0 Then NormalizeExpr = "": Exit Function

    ' تحويل الأرقام العربية-الهندية إلى أرقام لاتينية
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c)
        If code >= &H660 And code <= &H669 Then          ' ٠-٩ عربي
            sb = sb & Chr$(48 + (code - &H660))
        ElseIf code >= &H6F0 And code <= &H6F9 Then      ' ۰-۹ فارسي
            sb = sb & Chr$(48 + (code - &H6F0))
        ElseIf code = &H66B Then                          ' الفاصلة العشرية العربية
            sb = sb & "."
        ElseIf code = &H66C Then                          ' فاصلة الآلاف العربية
            ' تُحذف
        ElseIf code = &HD7 Then                           ' ×
            sb = sb & "*"
        ElseIf code = &HF7 Then                           ' ÷
            sb = sb & "/"
        ElseIf code = &H2212 Then                         ' − (سالب طويل)
            sb = sb & "-"
        ElseIf code = &H200F Or code = &H200E Or code = &H61C Then
            ' علامات اتجاه غير مرئية تُحذف
        Else
            sb = sb & c
        End If
    Next i

    sb = Trim$(sb)
    If Left$(sb, 1) = "=" Then sb = Mid$(sb, 2)
    sb = Replace(sb, "٫", ".")
    sb = Replace(sb, "،", ",")

    NormalizeExpr = "=" & sb
End Function

'--------------------------------------------------------------------
' عرض الأرقام بدقة عالية بدون ترميز علمي مزعج
'--------------------------------------------------------------------
Public Function NumToText(ByVal v As Variant) As String
    On Error GoTo Simple
    Dim s As String

    If VarType(v) = vbDouble Or VarType(v) = vbSingle Then
        ' تقريب تجميلي لـ 15 خانة معنوية — نفس اللي بيعمله إكسل عند العرض،
        ' عشان 1250*1.16 تطلع 1450 مش 1449.9999999999998
        s = NumStr(CosmeticRound(CDbl(v)))
    Else
        s = NumStr(v)
    End If

    NumToText = TidyNumber(s)
    Exit Function
Simple:
    NumToText = TidyNumber(Replace(CStr(v), ",", "."))
End Function

'--- تحويل رقم إلى نص بدون تأثّر بإعدادات اللغة (Str$ بيستخدم النقطة دايماً)
Private Function NumStr(ByVal v As Variant) As String
    On Error GoTo Fail
    NumStr = Trim$(Str$(v))
    Exit Function
Fail:
    NumStr = Trim$(Replace(CStr(v), ",", "."))
End Function

Private Function CosmeticRound(ByVal d As Double) As Double
    On Error GoTo Fail
    If d = 0 Then CosmeticRound = 0: Exit Function

    Dim a As Double, mag As Long, f As Double
    a = Abs(d)
    If a >= 1E+15 Or a <= 1E-14 Then CosmeticRound = d: Exit Function

    mag = Int(Log(a) / Log(10#))
    f = 10 ^ (14 - mag)
    If f <= 0 Then CosmeticRound = d: Exit Function

    CosmeticRound = Int(d * f + IIf(d < 0, -0.5, 0.5)) / f
    Exit Function
Fail:
    CosmeticRound = d
End Function

'--- تنظيف شكل الرقم: أصفار زائدة، وصفر قبل الفاصلة -----------------
Private Function TidyNumber(ByVal s As String) As String
    s = Trim$(Replace(s, ",", "."))

    If InStr(s, ".") > 0 And InStr(1, s, "E", vbTextCompare) = 0 Then
        Do While Right$(s, 1) = "0"
            s = Left$(s, Len(s) - 1)
        Loop
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    End If

    ' Str$ بيرجع ".3" بدل "0.3"
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    If s = "-0" Or s = "" Then s = "0"

    TidyNumber = s
End Function

Private Function ArrayToText(ByVal v As Variant) As String
    On Error GoTo Fail
    Dim r As Long, c As Long, sb As String
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    r1 = LBound(v, 1): r2 = UBound(v, 1)
    On Error Resume Next
    c1 = LBound(v, 2): c2 = UBound(v, 2)
    If Err.Number <> 0 Then
        Err.Clear
        For r = r1 To r2
            sb = sb & CStr(v(r)) & vbLf
        Next r
        ArrayToText = sb
        Exit Function
    End If
    On Error GoTo Fail
    If (r2 - r1) > 199 Then r2 = r1 + 199
    For r = r1 To r2
        For c = c1 To c2
            If c > c1 Then sb = sb & vbTab
            sb = sb & CStr(v(r, c))
        Next c
        sb = sb & vbLf
    Next r
    ArrayToText = sb
    Exit Function
Fail:
    ArrayToText = "(نتيجة مصفوفية)"
End Function

Private Function ErrName(ByVal v As Variant) As String
    Select Case CLng(v)
        Case 2000: ErrName = "#NULL!"
        Case 2007: ErrName = "#DIV/0!  (قسمة على صفر)"
        Case 2015: ErrName = "#VALUE!  (نوع قيمة غير مناسب)"
        Case 2023: ErrName = "#REF!  (مرجع خلية غير صالح)"
        Case 2029: ErrName = "#NAME?  (اسم دالة غير معروف)"
        Case 2036: ErrName = "#NUM!  (رقم خارج المدى)"
        Case 2042: ErrName = "#N/A  (القيمة غير موجودة)"
        Case Else: ErrName = "#ERROR"
    End Select
End Function

'====================================================================
' جمع مالي عالي الدقة (28 خانة) لتفادي أخطاء الفاصلة العائمة
'====================================================================
Public Function PreciseSum(ByVal rng As Object) As String
    On Error GoTo Fail
    Dim cel As Object
    Dim total As Variant
    total = CDec(0)
    Dim cnt As Long
    For Each cel In rng.Cells
        If IsNumeric(cel.Value) And Not IsEmpty(cel.Value) Then
            total = CDec(total) + CDec(cel.Value)
            cnt = cnt + 1
        End If
    Next cel
    PreciseSum = NumToText(total) & "  (عدد الخلايا الرقمية: " & cnt & ")"
    Exit Function
Fail:
    PreciseSum = "تعذّر الجمع الدقيق: " & Err.Description
End Function

'====================================================================
' فحص سلامة رقم مالي: هل فيه فروقات تقريب؟
'====================================================================
Public Function RoundingCheck(ByVal a As Variant, ByVal b As Variant, _
                              Optional ByVal tol As Double = 0.005) As String
    On Error GoTo Fail
    Dim diff As Variant
    diff = CDec(a) - CDec(b)
    If Abs(CDbl(diff)) <= tol Then
        RoundingCheck = "مطابق (الفرق " & NumToText(diff) & ")"
    Else
        RoundingCheck = "غير مطابق — الفرق = " & NumToText(diff)
    End If
    Exit Function
Fail:
    RoundingCheck = "تعذّر المقارنة."
End Function
