Attribute VB_Name = "modJson"
'====================================================================
' modJson  -  محلل ومُنشئ JSON بلغة VBA خالصة (بدون أي مراجع خارجية)
' جزء من إضافة: مساعد المحاسب (MaliAI)
'====================================================================
Option Explicit
Option Compare Binary

Private mS As String
Private mP As Long
Private mN As Long

'--------------------------------------------------------------------
' تحويل نص JSON إلى كائنات: Object(Scripting.Dictionary) أو Collection
' أو قيم بسيطة (String / Double / Boolean / Null)
'--------------------------------------------------------------------
Public Function JsonParse(ByVal sText As String) As Variant
    mS = sText
    mP = 1
    mN = Len(sText)
    Dim out As Variant
    SkipWs
    ParseValue out
    If IsObject(out) Then Set JsonParse = out Else JsonParse = out
End Function

'--- محاولة تحليل آمنة: ترجع Nothing عند الفشل بدل ما ترمي خطأ -------
Public Function JsonTryParse(ByVal sText As String) As Variant
    On Error GoTo Fail
    Dim v As Variant
    v = JsonParse(sText)
    If IsObject(v) Then Set JsonTryParse = v Else JsonTryParse = v
    Exit Function
Fail:
    Set JsonTryParse = Nothing
End Function

'====================================================================
' التحليل
'====================================================================
Private Sub ParseValue(ByRef out As Variant)
    SkipWs
    If mP > mN Then Err.Raise vbObjectError + 700, "modJson", "JSON: نهاية غير متوقعة للنص"

    Dim c As String
    c = Mid$(mS, mP, 1)

    Select Case c
        Case "{":  ParseObject out
        Case "[":  ParseArray out
        Case """": out = ParseString()
        Case "t":  Expect "true":  out = True
        Case "f":  Expect "false": out = False
        Case "n":  Expect "null":  out = Null
        Case Else: out = ParseNumber()
    End Select
End Sub

Private Sub Expect(ByVal lit As String)
    If LCase$(Mid$(mS, mP, Len(lit))) <> lit Then
        Err.Raise vbObjectError + 701, "modJson", "JSON: قيمة غير صالحة عند الموضع " & mP
    End If
    mP = mP + Len(lit)
End Sub

Private Sub ParseObject(ByRef out As Variant)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    mP = mP + 1                       ' تخطي {
    SkipWs
    If Mid$(mS, mP, 1) = "}" Then
        mP = mP + 1
        Set out = d
        Exit Sub
    End If

    Dim k As String, c As String
    Dim v As Variant
    Do
        SkipWs
        If Mid$(mS, mP, 1) <> """" Then
            Err.Raise vbObjectError + 702, "modJson", "JSON: اسم حقل غير صالح عند " & mP
        End If
        k = ParseString()
        SkipWs
        If Mid$(mS, mP, 1) <> ":" Then
            Err.Raise vbObjectError + 703, "modJson", "JSON: ':' مفقودة عند " & mP
        End If
        mP = mP + 1

        ParseValue v
        If IsObject(v) Then Set d.Item(k) = v Else d.Item(k) = v

        SkipWs
        c = Mid$(mS, mP, 1)
        mP = mP + 1
        If c = "}" Then Exit Do
        If c <> "," Then Err.Raise vbObjectError + 704, "modJson", "JSON: ',' مفقودة عند " & mP
    Loop

    Set out = d
End Sub

Private Sub ParseArray(ByRef out As Variant)
    Dim col As Collection
    Set col = New Collection

    mP = mP + 1                       ' تخطي [
    SkipWs
    If Mid$(mS, mP, 1) = "]" Then
        mP = mP + 1
        Set out = col
        Exit Sub
    End If

    Dim v As Variant, c As String
    Do
        ParseValue v
        col.Add v

        SkipWs
        c = Mid$(mS, mP, 1)
        mP = mP + 1
        If c = "]" Then Exit Do
        If c <> "," Then Err.Raise vbObjectError + 705, "modJson", "JSON: ',' مفقودة داخل مصفوفة عند " & mP
    Loop

    Set out = col
End Sub

Private Function ParseString() As String
    Dim sb As String
    Dim c As String, h As String

    mP = mP + 1                       ' تخطي علامة التنصيص الأولى
    Do While mP <= mN
        c = Mid$(mS, mP, 1)
        If c = """" Then
            mP = mP + 1
            ParseString = sb
            Exit Function
        ElseIf c = "\" Then
            mP = mP + 1
            c = Mid$(mS, mP, 1)
            Select Case c
                Case """": sb = sb & """"
                Case "\":  sb = sb & "\"
                Case "/":  sb = sb & "/"
                Case "b":  sb = sb & Chr$(8)
                Case "f":  sb = sb & Chr$(12)
                Case "n":  sb = sb & vbLf
                Case "r":  sb = sb & vbCr
                Case "t":  sb = sb & vbTab
                Case "u"
                    h = Mid$(mS, mP + 1, 4)
                    sb = sb & ChrW$(CLng("&H" & h))
                    mP = mP + 4
                Case Else
                    sb = sb & c
            End Select
            mP = mP + 1
        Else
            ' نسخ دفعة واحدة لتسريع الأداء على النصوص الطويلة
            Dim st As Long, q As Long, b As Long
            st = mP
            q = InStr(mP, mS, """")
            b = InStr(mP, mS, "\")
            If q = 0 Then q = mN + 1
            If b = 0 Then b = mN + 1
            Dim stopAt As Long
            stopAt = IIf(q < b, q, b)
            sb = sb & Mid$(mS, st, stopAt - st)
            mP = stopAt
        End If
    Loop
    Err.Raise vbObjectError + 706, "modJson", "JSON: نص غير منتهٍ"
End Function

Private Function ParseNumber() As Variant
    Dim st As Long, c As String
    st = mP
    Do While mP <= mN
        c = Mid$(mS, mP, 1)
        If InStr("0123456789+-.eE", c) = 0 Then Exit Do
        mP = mP + 1
    Loop
    If mP = st Then Err.Raise vbObjectError + 707, "modJson", "JSON: رقم غير صالح عند " & st
    ' Val لا يتأثر بإعدادات اللغة (يستخدم النقطة دائماً)
    ParseNumber = Val(Mid$(mS, st, mP - st))
End Function

Private Sub SkipWs()
    Dim c As String
    Do While mP <= mN
        c = Mid$(mS, mP, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
        mP = mP + 1
    Loop
End Sub

'====================================================================
' الإنشاء (Serialization)
'====================================================================

'--- تهريب كامل: كل حرف غير ASCII يتحول إلى \uXXXX -------------------
'    هيك جسم الطلب بيصير ASCII صافي فما بصير مشاكل ترميز مع العربي
Public Function JsonEscape(ByVal s As String) As String
    Dim i As Long, n As Long, code As Long, ch As String
    Dim sb As String
    n = Len(s)
    For i = 1 To n
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34:  sb = sb & "\"""
            Case 92:  sb = sb & "\\"
            Case 8:   sb = sb & "\b"
            Case 9:   sb = sb & "\t"
            Case 10:  sb = sb & "\n"
            Case 12:  sb = sb & "\f"
            Case 13:  sb = sb & "\r"
            Case Else
                If code < 32 Or code > 126 Then
                    sb = sb & "\u" & Right$("000" & Hex$(code), 4)
                Else
                    sb = sb & ch
                End If
        End Select
    Next i
    JsonEscape = sb
End Function

Public Function JsonStr(ByVal s As String) As String
    JsonStr = """" & JsonEscape(s) & """"
End Function

Public Function JsonNum(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ",", ".")           ' حماية من الفاصلة العشرية العربية
    JsonNum = s
End Function

'====================================================================
' مساعدات الوصول للقيم:  JGet(obj, "choices/0/message/content")
'====================================================================
Public Function JGet(ByVal root As Variant, ByVal path As String) As Variant
    On Error GoTo Fail
    Dim parts() As String, i As Long
    Dim cur As Variant

    If IsObject(root) Then Set cur = root Else cur = root
    parts = Split(path, "/")

    For i = 0 To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo Continue_
        If Not IsObject(cur) Then GoTo Fail

        If TypeName(cur) = "Dictionary" Then
            If Not cur.Exists(parts(i)) Then GoTo Fail
            If IsObject(cur.Item(parts(i))) Then
                Set cur = cur.Item(parts(i))
            Else
                cur = cur.Item(parts(i))
            End If
        ElseIf TypeName(cur) = "Collection" Then
            Dim idx As Long
            idx = CLng(parts(i)) + 1        ' فهرسة من الصفر في المسار
            If idx < 1 Or idx > cur.Count Then GoTo Fail
            If IsObject(cur.Item(idx)) Then
                Set cur = cur.Item(idx)
            Else
                cur = cur.Item(idx)
            End If
        Else
            GoTo Fail
        End If
Continue_:
    Next i

    If IsObject(cur) Then Set JGet = cur Else JGet = cur
    Exit Function
Fail:
    JGet = Empty
End Function

'--- نفس السابق لكن يرجع نص دائماً ----------------------------------
Public Function JGetStr(ByVal root As Variant, ByVal path As String) As String
    Dim v As Variant
    v = JGet(root, path)
    If IsObject(v) Then
        JGetStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        JGetStr = ""
    Else
        JGetStr = CStr(v)
    End If
End Function

Public Function JHas(ByVal root As Variant, ByVal path As String) As Boolean
    Dim v As Variant
    v = JGet(root, path)
    JHas = Not (IsEmpty(v))
End Function
