Attribute VB_Name = "modApi"
'====================================================================
' modApi  -  طبقة موحّدة لمزوّدي الذكاء الاصطناعي
'            NVIDIA | OpenRouter | Groq | Anthropic | مخصّص
'
' النسخة 2: استدعاء الأدوات الأصلي (Function Calling).
' النموذج ما بيكتب JSON بإيده — المزوّد بيرجّع الاستدعاء منظّم.
'====================================================================
Option Explicit

Private Const NV_BASE As String = "https://integrate.api.nvidia.com/v1"
Private Const OR_BASE As String = "https://openrouter.ai/api/v1"
Private Const GQ_BASE As String = "https://api.groq.com/openai/v1"
Private Const AN_BASE As String = "https://api.anthropic.com/v1"
Private Const AN_VERSION As String = "2023-06-01"

Public Type AiReply
    Ok As Boolean
    Text As String
    ErrText As String
    ToolCalls As Object          ' Collection من Dictionary: id / name / args
    RawToolCalls As String       ' JSON الخام لإعادة إرساله كما هو
End Type

'====================================================================
' بناء سجل الرسائل
'====================================================================
Public Function NewHistory() As Collection
    Set NewHistory = New Collection
End Function

Private Function NewMsg() As Object
    Set NewMsg = CreateObject("Scripting.Dictionary")
End Function

Public Sub AddMsg(ByVal hist As Collection, ByVal role As String, ByVal content As String)
    Dim d As Object
    Set d = NewMsg()
    d.Item("role") = role
    d.Item("content") = content
    hist.Add d
End Sub

'--- رسالة المساعد اللي فيها استدعاءات أدوات (تُعاد كما هي) ----------
Public Sub AddAssistantToolCalls(ByVal hist As Collection, ByVal rawToolCalls As String, _
                                 ByVal content As String)
    Dim d As Object
    Set d = NewMsg()
    d.Item("role") = "assistant"
    d.Item("content") = content
    d.Item("tool_calls") = rawToolCalls
    hist.Add d
End Sub

'--- نتيجة تنفيذ أداة ------------------------------------------------
Public Sub AddToolResult(ByVal hist As Collection, ByVal toolCallId As String, _
                         ByVal content As String)
    Dim d As Object
    Set d = NewMsg()
    d.Item("role") = "tool"
    d.Item("tool_call_id") = toolCallId
    d.Item("content") = content
    hist.Add d
End Sub

'--- قص السجل مع الحفاظ على تماسك أزواج (استدعاء/نتيجة) -------------
Public Sub TrimHistory(ByVal hist As Collection, Optional ByVal keepLast As Long = 40)
    Do While hist.Count > keepLast
        ' ما منقص رسالة tool لحالها بدون الـ assistant تبعتها
        If hist.Item(1).Item("role") = "tool" And hist.Count > 1 Then
            hist.Remove 1
        Else
            hist.Remove 1
        End If
        ' لو أول رسالة صارت tool يتيمة، منشيلها كمان
        Do While hist.Count > 0
            If hist.Item(1).Item("role") = "tool" Then hist.Remove 1 Else Exit Do
        Loop
    Loop
End Sub

'====================================================================
' هل المزوّد بيدعم استدعاء الأدوات الأصلي؟
'====================================================================
Public Function SupportsNativeTools(Optional ByVal provider As String = "") As Boolean
    If Len(provider) = 0 Then provider = CfgProvider()
    Select Case LCase$(provider)
        Case "omniroute", "nvidia", "openrouter", "groq", "anthropic"
            SupportsNativeTools = True
        Case "custom"
            SupportsNativeTools = CfgNativeTools()
        Case Else
            SupportsNativeTools = False
    End Select
End Function

Public Function BaseUrlFor(ByVal provider As String) As String
    Select Case LCase$(provider)
        Case "omniroute":  BaseUrlFor = OMNI_BASE
        Case "nvidia":     BaseUrlFor = NV_BASE
        Case "openrouter": BaseUrlFor = OR_BASE
        Case "groq":       BaseUrlFor = GQ_BASE
        Case "anthropic":  BaseUrlFor = AN_BASE
        Case "custom":     BaseUrlFor = CfgBaseUrl()
        Case Else:         BaseUrlFor = NV_BASE
    End Select
End Function

'====================================================================
' الاستدعاء الرئيسي مع إعادة محاولة تلقائية
'====================================================================
Public Function AiChat(ByVal systemPrompt As String, ByVal hist As Collection, _
                       Optional ByVal toolsJson As String = "") As AiReply
    Dim attempt As Long, r As AiReply
    For attempt = 1 To 3
        r = AiChatOnce(systemPrompt, hist, toolsJson)
        If r.Ok Then AiChat = r: Exit Function
        If Not IsTransient(r.ErrText) Then AiChat = r: Exit Function
        On Error Resume Next
        Application.Wait Now + TimeSerial(0, 0, 2)
        On Error GoTo 0
    Next attempt
    AiChat = r
End Function

Private Function IsTransient(ByVal e As String) As Boolean
    If InStr(e, "رد غير مفهوم") > 0 Then IsTransient = True: Exit Function
    If InStr(e, "خلل مؤقت") > 0 Then IsTransient = True: Exit Function
    If InStr(e, "429") > 0 Then IsTransient = True: Exit Function
    If InStr(e, "انتهت مهلة") > 0 Then IsTransient = True: Exit Function
    If InStr(e, "رداً فارغاً") > 0 Then IsTransient = True: Exit Function
End Function

Private Function AiChatOnce(ByVal systemPrompt As String, ByVal hist As Collection, _
                            ByVal toolsJson As String) As AiReply
    Dim p As String
    p = CfgProvider()
    If p = "anthropic" Or (p = "custom" And CfgApiFormat() = "anthropic") Then
        AiChatOnce = CallAnthropic(p, systemPrompt, hist, toolsJson)
    Else
        AiChatOnce = CallOpenAiCompatible(p, systemPrompt, hist, toolsJson)
    End If
End Function

'====================================================================
' تسلسل سجل الرسائل إلى JSON بصيغة OpenAI
'====================================================================
Private Function SerializeMessages(ByVal systemPrompt As String, ByVal hist As Collection) As String
    Dim sb As String, i As Long, d As Object
    sb = "{""role"":""system"",""content"":" & JsonStr(systemPrompt) & "}"

    For i = 1 To hist.Count
        Set d = hist.Item(i)
        sb = sb & ","
        Select Case CStr(d.Item("role"))

            Case "tool"
                sb = sb & "{""role"":""tool"",""tool_call_id"":" & _
                     JsonStr(CStr(d.Item("tool_call_id"))) & _
                     ",""content"":" & JsonStr(CStr(d.Item("content"))) & "}"

            Case "assistant"
                If d.Exists("tool_calls") Then
                    sb = sb & "{""role"":""assistant"",""content"":" & _
                         IIf(Len(CStr(d.Item("content"))) > 0, JsonStr(CStr(d.Item("content"))), "null") & _
                         ",""tool_calls"":" & CStr(d.Item("tool_calls")) & "}"
                Else
                    sb = sb & "{""role"":""assistant"",""content"":" & _
                         JsonStr(CStr(d.Item("content"))) & "}"
                End If

            Case Else
                sb = sb & "{""role"":" & JsonStr(CStr(d.Item("role"))) & _
                     ",""content"":" & JsonStr(CStr(d.Item("content"))) & "}"
        End Select
    Next i

    SerializeMessages = sb
End Function

'====================================================================
' NVIDIA / OpenRouter / Groq / مخصّص  (صيغة OpenAI)
'====================================================================
Private Function CallOpenAiCompatible(ByVal provider As String, ByVal systemPrompt As String, _
                                      ByVal hist As Collection, ByVal toolsJson As String) As AiReply
    Dim r As AiReply
    Dim baseUrl As String, key As String, sb As String

    key = CfgApiKey(provider)
    If Len(key) < 5 And LCase$(provider) <> "omniroute" Then
        r.ErrText = "ما في مفتاح محفوظ للمزوّد " & ProviderTitle(provider) & ". افتح الإعدادات وضيفه."
        CallOpenAiCompatible = r
        Exit Function
    End If

    baseUrl = BaseUrlFor(provider)
    If Len(baseUrl) = 0 Then
        r.ErrText = "ما في رابط endpoint محفوظ. افتح الإعدادات وحطّه."
        CallOpenAiCompatible = r
        Exit Function
    End If

    sb = "{""model"":" & JsonStr(CfgModel(provider))
    sb = sb & ",""temperature"":" & JsonNum(CfgTemperature())
    sb = sb & ",""max_tokens"":4096"
    If Len(toolsJson) > 0 Then
        sb = sb & ",""tools"":" & toolsJson & ",""tool_choice"":""auto"""
    End If
    sb = sb & ",""messages"":[" & SerializeMessages(systemPrompt, hist) & "]}"

    Dim hdrs As Variant
    If provider = "openrouter" Then
        hdrs = Array(Array("Authorization", "Bearer " & key), _
                     Array("HTTP-Referer", "https://local.excel.addin/maliai"), _
                     Array("X-Title", "MaliAI Excel Assistant"))
    ElseIf Len(key) > 0 Then
        hdrs = Array(Array("Authorization", "Bearer " & key))
    Else
        hdrs = Array()
    End If

    Dim res As HttpResult
    res = HttpPostJson(baseUrl & "/chat/completions", sb, hdrs)

    If Not res.Ok Then
        r.ErrText = FriendlyError(res, provider)
        CallOpenAiCompatible = r
        Exit Function
    End If

    Dim j As Variant
    Set j = JsonTryParse(res.Body)
    If j Is Nothing Then
        r.ErrText = "رد غير مفهوم من الخادم:" & vbCrLf & Left$(res.Body, 400)
        CallOpenAiCompatible = r
        Exit Function
    End If

    Dim msg As Variant
    Set msg = JGet(j, "choices/0/message")
    If msg Is Nothing Then
        Dim apiErr As String
        apiErr = JGetStr(j, "error/message")
        r.ErrText = IIf(Len(apiErr) > 0, "الخادم: " & apiErr, "النموذج رجّع رداً فارغاً.")
        CallOpenAiCompatible = r
        Exit Function
    End If

    r.Text = JGetStr(msg, "content")

    ' --- استخراج استدعاءات الأدوات ---
    Set r.ToolCalls = New Collection
    Dim tc As Variant
    Set tc = JGet(msg, "tool_calls")
    If Not tc Is Nothing Then
        If TypeName(tc) = "Collection" Then
            Dim i As Long, raw As String
            For i = 1 To tc.Count
                Dim c As Object
                Set c = CreateObject("Scripting.Dictionary")
                c.Item("id") = JGetStr(tc.Item(i), "id")
                c.Item("name") = JGetStr(tc.Item(i), "function/name")
                c.Item("args") = JGetStr(tc.Item(i), "function/arguments")
                r.ToolCalls.Add c

                If i > 1 Then raw = raw & ","
                raw = raw & "{""id"":" & JsonStr(CStr(c.Item("id"))) & _
                      ",""type"":""function"",""function"":{""name"":" & JsonStr(CStr(c.Item("name"))) & _
                      ",""arguments"":" & JsonStr(CStr(c.Item("args"))) & "}}"
            Next i
            If Len(raw) > 0 Then r.RawToolCalls = "[" & raw & "]"
        End If
    End If

    If Len(r.Text) = 0 And r.ToolCalls.Count = 0 Then
        r.ErrText = "النموذج رجّع رداً فارغاً. جرّب نموذجاً آخر من الإعدادات."
        CallOpenAiCompatible = r
        Exit Function
    End If

    r.Ok = True
    CallOpenAiCompatible = r
End Function

'====================================================================
' Anthropic  (صيغة مختلفة للأدوات)
'====================================================================
Private Function CallAnthropic(ByVal provider As String, ByVal systemPrompt As String, _
                               ByVal hist As Collection, ByVal toolsJson As String) As AiReply
    Dim r As AiReply
    Dim key As String, baseUrl As String, sb As String, i As Long

    key = CfgApiKey(provider)
    If Len(key) < 5 Then
        r.ErrText = "ما في مفتاح محفوظ للمزوّد " & ProviderTitle(provider) & "."
        CallAnthropic = r
        Exit Function
    End If
    baseUrl = BaseUrlFor(provider)

    sb = "{""model"":" & JsonStr(CfgModel(provider))
    sb = sb & ",""max_tokens"":4096"
    sb = sb & ",""temperature"":" & JsonNum(CfgTemperature())
    sb = sb & ",""system"":" & JsonStr(systemPrompt)
    If Len(toolsJson) > 0 Then
        sb = sb & ",""tools"":" & AnthropicTools(toolsJson)
    End If
    sb = sb & ",""messages"":["

    Dim first As Boolean
    first = True
    For i = 1 To hist.Count
        Dim d As Object
        Set d = hist.Item(i)
        Dim role As String
        role = CStr(d.Item("role"))
        If role = "tool" Then role = "user"
        If Not first Then sb = sb & ","
        sb = sb & "{""role"":" & JsonStr(role) & ",""content"":" & _
             JsonStr(CStr(d.Item("content"))) & "}"
        first = False
    Next i
    sb = sb & "]}"

    Dim hdrs As Variant
    hdrs = Array(Array("x-api-key", key), _
                 Array("Authorization", "Bearer " & key), _
                 Array("anthropic-version", AN_VERSION))

    Dim res As HttpResult
    res = HttpPostJson(baseUrl & "/messages", sb, hdrs)

    If Not res.Ok Then
        r.ErrText = FriendlyError(res, provider)
        CallAnthropic = r
        Exit Function
    End If

    Dim j As Variant
    Set j = JsonTryParse(res.Body)
    If j Is Nothing Then
        r.ErrText = "رد غير مفهوم من الخادم:" & vbCrLf & Left$(res.Body, 400)
        CallAnthropic = r
        Exit Function
    End If

    Set r.ToolCalls = New Collection
    Dim txt As String
    Dim blocks As Variant
    Set blocks = JGet(j, "content")
    If Not blocks Is Nothing Then
        If TypeName(blocks) = "Collection" Then
            For i = 1 To blocks.Count
                If JGetStr(blocks.Item(i), "type") = "text" Then
                    txt = txt & JGetStr(blocks.Item(i), "text")
                End If
            Next i
        End If
    End If
    If Len(txt) = 0 Then txt = JGetStr(j, "choices/0/message/content")

    If Len(txt) = 0 Then
        Dim apiErr As String
        apiErr = JGetStr(j, "error/message")
        r.ErrText = IIf(Len(apiErr) > 0, "الخادم: " & apiErr, "رد فارغ من النموذج.")
        CallAnthropic = r
        Exit Function
    End If

    r.Ok = True
    r.Text = txt
    CallAnthropic = r
End Function

'--- تحويل تعريف أدوات OpenAI إلى صيغة Anthropic ---------------------
Private Function AnthropicTools(ByVal toolsJson As String) As String
    On Error GoTo Fail
    Dim arr As Variant
    Set arr = JsonTryParse(toolsJson)
    If arr Is Nothing Then AnthropicTools = "[]": Exit Function

    Dim sb As String, i As Long
    For i = 1 To arr.Count
        If i > 1 Then sb = sb & ","
        sb = sb & "{""name"":" & JsonStr(JGetStr(arr.Item(i), "function/name")) & _
             ",""description"":" & JsonStr(JGetStr(arr.Item(i), "function/description")) & _
             ",""input_schema"":{""type"":""object"",""properties"":{}}}"
    Next i
    AnthropicTools = "[" & sb & "]"
    Exit Function
Fail:
    AnthropicTools = "[]"
End Function

'====================================================================
' رسائل خطأ مفهومة بالعربي
'====================================================================
Private Function FriendlyError(ByRef res As HttpResult, ByVal provider As String) As String
    Dim detail As String
    Dim j As Variant
    Set j = JsonTryParse(res.Body)
    If Not j Is Nothing Then
        detail = JGetStr(j, "error/message")
        If Len(detail) = 0 Then detail = JGetStr(j, "message")
        If Len(detail) = 0 Then detail = JGetStr(j, "detail")
    End If

    Select Case res.Status
        Case 401, 403
            FriendlyError = "المفتاح مرفوض أو منتهي (" & res.Status & "). تأكد من مفتاح " & _
                            ProviderTitle(provider) & " في الإعدادات."
        Case 402
            FriendlyError = "الرصيد غير كافٍ عند " & ProviderTitle(provider) & "."
        Case 404
            FriendlyError = "اسم النموذج أو الرابط غير موجود (404). اضغط ""تحديث النماذج"" بالإعدادات."
        Case 422, 400
            FriendlyError = "الطلب مرفوض (" & res.Status & "). غالباً النموذج ما بيدعم استدعاء الأدوات — " & _
                            "جرّب نموذج تاني، أو عطّل الأدوات الأصلية من الإعدادات."
        Case 429
            FriendlyError = "تجاوزت حد الطلبات (429). استنّى دقيقة أو بدّل النموذج."
        Case 500 To 599
            FriendlyError = "خلل مؤقت عند المزوّد (" & res.Status & "). جرّب بعد شوي."
        Case 0
            If LCase$(provider) = "omniroute" Then
                FriendlyError = "OmniRoute مش شغّالة على الجهاز." & vbCrLf & _
                    "افتح موجّه الأوامر (cmd) واكتب:   omniroute" & vbCrLf & _
                    "وإذا أول مرة، ركّبها:   npm i -g omniroute"
            Else
                FriendlyError = res.ErrText
            End If
        Case Else
            FriendlyError = "فشل الطلب (" & res.Status & ")."
    End Select

    If Len(detail) > 0 Then FriendlyError = FriendlyError & vbCrLf & "التفاصيل: " & Left$(detail, 300)
End Function

'====================================================================
' جلب قائمة النماذج
'====================================================================
Public Function ListModels(ByVal provider As String, ByRef errMsg As String) As Collection
    Dim out As New Collection
    Dim key As String, bu As String
    key = CfgApiKey(provider)
    bu = BaseUrlFor(provider)
    errMsg = ""

    If Len(bu) = 0 Then errMsg = "حط رابط الـ endpoint أولاً.": Set ListModels = out: Exit Function

    Dim res As HttpResult
    Select Case LCase$(provider)
        Case "openrouter"
            res = HttpGet(bu & "/models", Array())
        Case "anthropic"
            If Len(key) < 5 Then errMsg = "لازم المفتاح أولاً.": Set ListModels = out: Exit Function
            res = HttpGet(bu & "/models?limit=100", _
                          Array(Array("x-api-key", key), Array("anthropic-version", AN_VERSION)))
        Case "omniroute"
            res = HttpGet(bu & "/models", Array())
            If Not res.Ok Then
                errMsg = "ما قدرت أوصل OmniRoute على " & bu & vbCrLf & _
                         "تأكد إنها شغّالة: افتح موجّه الأوامر واكتب  omniroute  " & _
                         "(أو ركّبها أول مرة بـ  npm i -g omniroute )"
                Set ListModels = out
                Exit Function
            End If
        Case Else
            If Len(key) < 5 Then errMsg = "لازم المفتاح أولاً.": Set ListModels = out: Exit Function
            res = HttpGet(bu & "/models", Array(Array("Authorization", "Bearer " & key)))
    End Select

    If Not res.Ok Then
        errMsg = FriendlyError(res, provider)
        Set ListModels = out
        Exit Function
    End If

    Dim j As Variant
    Set j = JsonTryParse(res.Body)
    If j Is Nothing Then errMsg = "تعذّر قراءة قائمة النماذج.": Set ListModels = out: Exit Function

    Dim arr As Variant
    Set arr = JGet(j, "data")
    If arr Is Nothing Then Set ListModels = out: Exit Function
    If TypeName(arr) <> "Collection" Then Set ListModels = out: Exit Function

    Dim i As Long, id As String, nm As String, tag As String, isFree As Boolean
    For i = 1 To arr.Count
        id = JGetStr(arr.Item(i), "id")
        If Len(id) > 0 Then
            nm = JGetStr(arr.Item(i), "name")
            If Len(nm) = 0 Then nm = JGetStr(arr.Item(i), "display_name")

            isFree = (InStr(1, id, ":free", vbTextCompare) > 0)
            If LCase$(provider) = "openrouter" Then
                Dim pp As String
                pp = JGetStr(arr.Item(i), "pricing/prompt")
                If Len(pp) > 0 Then If Val(pp) = 0 Then isFree = True
            ElseIf LCase$(provider) = "groq" Or LCase$(provider) = "nvidia" _
                Or LCase$(provider) = "omniroute" Then
                isFree = True
            End If
            tag = IIf(isFree, "  ★ مجاني", "")

            If (Not CfgFreeOnly()) Or isFree Then out.Add id & "|" & nm & tag
        End If
    Next i

    If out.Count = 0 Then errMsg = "ما لقيت نماذج مطابقة. ألغِ خيار ""المجاني فقط""."
    Set ListModels = out
End Function

'====================================================================
' اختبار سريع للاتصال
'====================================================================
Public Function TestConnection() As String
    Dim h As Collection
    Set h = NewHistory()
    AddMsg h, "user", "رد بكلمة واحدة فقط: تمام"

    Dim rep As AiReply
    rep = AiChat("أنت مساعد. رد بكلمة واحدة.", h, "")
    TestConnection = IIf(rep.Ok, "", rep.ErrText)
End Function

'--- اختبار إن النموذج بيدعم استدعاء الأدوات فعلياً ------------------
Public Function TestToolSupport() As String
    Dim h As Collection
    Set h = NewHistory()
    AddMsg h, "user", "اقرأ النطاق A1:B5 من الورقة النشطة."

    Dim probe As String
    probe = "[{""type"":""function"",""function"":{""name"":""read_range""," & _
            """description"":""قراءة نطاق""," & _
            """parameters"":{""type"":""object"",""properties"":{""address"":{""type"":""string""}}," & _
            """required"":[""address""]}}}]"

    Dim rep As AiReply
    rep = AiChat("أنت مساعد إكسل. استخدم الأدوات المتاحة.", h, probe)

    If Not rep.Ok Then
        TestToolSupport = "✖ " & rep.ErrText
    ElseIf rep.ToolCalls Is Nothing Then
        TestToolSupport = "⚠ النموذج ما استدعى أداة — رح نستخدم الوضع النصي الاحتياطي."
    ElseIf rep.ToolCalls.Count = 0 Then
        TestToolSupport = "⚠ النموذج ما استدعى أداة — رح نستخدم الوضع النصي الاحتياطي."
    Else
        TestToolSupport = ""
    End If
End Function
