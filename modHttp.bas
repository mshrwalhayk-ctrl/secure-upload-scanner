Attribute VB_Name = "modHttp"
'====================================================================
' modHttp  -  طبقة الاتصال بالإنترنت (HTTPS) بترميز UTF-8 سليم
'====================================================================
Option Explicit

Public Type HttpResult
    Ok As Boolean
    Status As Long
    Body As String
    ErrText As String
End Type

Private Const TIMEOUT_RESOLVE As Long = 10000
Private Const TIMEOUT_CONNECT As Long = 15000
Private Const TIMEOUT_SEND As Long = 20000
Private Const TIMEOUT_RECEIVE As Long = 180000   ' 3 دقائق للطلبات الطويلة

'--------------------------------------------------------------------
' طلب POST بجسم JSON.  headers = مصفوفة ثنائية: Array(Array(name,value), ...)
'--------------------------------------------------------------------
Public Function HttpPostJson(ByVal url As String, ByVal body As String, _
                             ByVal headers As Variant) As HttpResult
    HttpPostJson = HttpSend("POST", url, body, headers)
End Function

Public Function HttpGet(ByVal url As String, ByVal headers As Variant) As HttpResult
    HttpGet = HttpSend("GET", url, "", headers)
End Function

'--------------------------------------------------------------------
Private Function HttpSend(ByVal method As String, ByVal url As String, _
                          ByVal body As String, ByVal headers As Variant) As HttpResult
    Dim r As HttpResult
    Dim http As Object
    Dim i As Long

    On Error GoTo Fail

    Set http = CreateHttp()
    If http Is Nothing Then
        r.Ok = False
        r.ErrText = "تعذّر تهيئة مكوّن الاتصال بالإنترنت (MSXML) على هذا الجهاز."
        HttpSend = r
        Exit Function
    End If

    On Error Resume Next
    http.setTimeouts TIMEOUT_RESOLVE, TIMEOUT_CONNECT, TIMEOUT_SEND, TIMEOUT_RECEIVE
    On Error GoTo Fail

    http.Open method, url, False

    If Len(body) > 0 Then
        http.setRequestHeader "Content-Type", "application/json"
    End If
    http.setRequestHeader "Accept", "application/json"

    If IsArray(headers) Then
        For i = LBound(headers) To UBound(headers)
            If IsArray(headers(i)) Then
                http.setRequestHeader CStr(headers(i)(0)), CStr(headers(i)(1))
            End If
        Next i
    End If

    If Len(body) > 0 Then
        ' الجسم مهرَّب بالكامل إلى ASCII في modJson، فالإرسال النصي آمن
        http.send body
    Else
        http.send
    End If

    r.Status = http.Status
    r.Body = ReadBodyUtf8(http)
    r.Ok = (r.Status >= 200 And r.Status < 300)
    If Not r.Ok Then
        r.ErrText = "الخادم رجّع رمز " & r.Status & vbCrLf & Left$(r.Body, 900)
    End If

    HttpSend = r
    Exit Function

Fail:
    r.Ok = False
    r.Status = 0
    Select Case Err.Number
        Case -2147012894, -2147012891
            r.ErrText = "انتهت مهلة الاتصال. تأكد من الإنترنت أو جرّب نموذجاً أسرع."
        Case -2147012867, -2147012865
            r.ErrText = "تعذّر الاتصال بالخادم. تحقق من الإنترنت أو من إعدادات البروكسي/الجدار الناري."
        Case Else
            r.ErrText = "خطأ في الاتصال (" & Err.Number & "): " & Err.Description
    End Select
    HttpSend = r
End Function

'--------------------------------------------------------------------
Private Function CreateHttp() As Object
    On Error Resume Next
    Set CreateHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If CreateHttp Is Nothing Then Set CreateHttp = CreateObject("MSXML2.ServerXMLHTTP")
    If CreateHttp Is Nothing Then Set CreateHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    If CreateHttp Is Nothing Then Set CreateHttp = CreateObject("MSXML2.XMLHTTP")
    Err.Clear
End Function

'--- فك ترميز الرد كـ UTF-8 مهما كانت ترويسة الخادم -----------------
Private Function ReadBodyUtf8(ByVal http As Object) As String
    On Error GoTo Fallback
    Dim b() As Byte
    b = http.responseBody
    If (Not Not b) = 0 Then GoTo Fallback   ' مصفوفة فارغة

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                              ' adTypeBinary
    st.Open
    st.Write b
    st.Position = 0
    st.Type = 2                              ' adTypeText
    st.Charset = "utf-8"
    ReadBodyUtf8 = st.ReadText
    st.Close
    Exit Function
Fallback:
    On Error Resume Next
    ReadBodyUtf8 = http.responseText
End Function

'--------------------------------------------------------------------
' ترميز نص لاستعماله داخل عنوان URL
'--------------------------------------------------------------------
Public Function UrlEncode(ByVal s As String) As String
    Dim i As Long, c As String, code As Long, sb As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) _
           Or (code >= 97 And code <= 122) Or InStr("-_.~", c) > 0 Then
            sb = sb & c
        Else
            Dim bts() As Byte, j As Long
            bts = Utf8Bytes(c)
            For j = LBound(bts) To UBound(bts)
                sb = sb & "%" & Right$("0" & Hex$(bts(j)), 2)
            Next j
        End If
    Next i
    UrlEncode = sb
End Function

Private Function Utf8Bytes(ByVal s As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1
    st.Position = 3                          ' تخطي BOM
    Utf8Bytes = st.Read
    st.Close
End Function
