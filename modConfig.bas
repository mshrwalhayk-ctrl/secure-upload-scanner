Attribute VB_Name = "modConfig"
'====================================================================
' modConfig  -  إعدادات الإضافة (المزوّد، المفتاح، النموذج، الأمان)
' تُحفظ في سجل ويندوز الخاص بالمستخدم الحالي فقط
'====================================================================
Option Explicit

Public Const APP_NAME As String = "MaliAI"
Public Const APP_TITLE As String = "مساعد المحاسب"
Public Const APP_VERSION As String = "1.0"

Private Const REG_APP As String = "MaliAI"
Private Const REG_SEC As String = "Settings"

'--------------------------------------------------------------------
' الإعداد الجاهز للمزوّد المخصّص — معبّى مسبقاً حتى المحاسب ما يحتاج
' يعمل شي غير إنه يلصق المفتاح ويضغط حفظ.
'--------------------------------------------------------------------
'  الافتراضي: OmniRoute — بوابة محلية بتوحّد كل المزوّدين، بتضغط النص،
'  وبتشتغل بدون أي مفتاح. تركيبها:  npm i -g omniroute
Public Const DEFAULT_PROVIDER As String = "omniroute"
Public Const OMNI_BASE        As String = "http://localhost:20128/v1"
Public Const OMNI_MODEL       As String = "auto"

Public Const PRESET_BASE_URL  As String = _
    "https://preview-chat-8c74296f-a7ef-49fc-ab00-41bf8b435603.space-z.ai/api/provider/v1"
Public Const PRESET_FORMAT    As String = "openai"
Public Const PRESET_MODEL     As String = "glm-5.3-flash"

'====================================================================
' قراءة / كتابة عامة
'====================================================================
Public Function CfgGet(ByVal key As String, Optional ByVal dflt As String = "") As String
    On Error Resume Next
    CfgGet = GetSetting(REG_APP, REG_SEC, key, dflt)
    If Err.Number <> 0 Then CfgGet = dflt: Err.Clear
End Function

Public Sub CfgSet(ByVal key As String, ByVal value As String)
    On Error Resume Next
    SaveSetting REG_APP, REG_SEC, key, value
End Sub

'====================================================================
' المزوّد:  openrouter | groq | anthropic | custom
'====================================================================
Public Function CfgProvider() As String
    Dim p As String
    p = LCase$(CfgGet("Provider", DEFAULT_PROVIDER))
    Select Case p
        Case "omniroute", "nvidia", "anthropic", "openrouter", "groq", "custom"
            CfgProvider = p
        Case Else
            CfgProvider = DEFAULT_PROVIDER
    End Select
End Function

Public Function ProviderTitle(ByVal p As String) As String
    Select Case LCase$(p)
        Case "omniroute":  ProviderTitle = "OmniRoute — بوابة محلية ★ (بدون مفتاح + ضغط نص)"
        Case "nvidia":     ProviderTitle = "NVIDIA NIM (نماذج قوية مجانية ★)"
        Case "openrouter": ProviderTitle = "OpenRouter (نماذج مجانية ★)"
        Case "groq":       ProviderTitle = "Groq (نماذج مجانية سريعة)"
        Case "anthropic":  ProviderTitle = "Anthropic (كلود مباشرة — مدفوع)"
        Case "custom":     ProviderTitle = "مزوّد مخصّص (رابط خاص)"
        Case Else:         ProviderTitle = p
    End Select
End Function

'====================================================================
' إعدادات المزوّد المخصّص
'   BaseUrl مثال:  https://api.example.com/v1        (نمط OpenAI)
'                  https://api.example.com/anthropic (نمط Anthropic)
'   ApiFormat:  openai | anthropic
'====================================================================
Public Function CfgBaseUrl() As String
    Dim u As String
    u = Trim$(CfgGet("BaseUrl_custom", PRESET_BASE_URL))
    Do While Right$(u, 1) = "/"
        u = Left$(u, Len(u) - 1)
    Loop
    CfgBaseUrl = u
End Function

Public Sub CfgSetBaseUrl(ByVal u As String)
    CfgSet "BaseUrl_custom", Trim$(u)
End Sub

Public Function CfgApiFormat() As String
    Dim f As String
    f = LCase$(CfgGet("ApiFormat_custom", PRESET_FORMAT))
    If f <> "anthropic" Then f = "openai"
    CfgApiFormat = f
End Function

Public Sub CfgSetApiFormat(ByVal f As String)
    CfgSet "ApiFormat_custom", LCase$(Trim$(f))
End Sub

'--- عرض النماذج المجانية فقط في قائمة الاختيار ----------------------
Public Function CfgFreeOnly() As Boolean
    CfgFreeOnly = (CfgGet("FreeOnly", "1") = "1")
End Function

Public Sub CfgSetFreeOnly(ByVal b As Boolean)
    CfgSet "FreeOnly", IIf(b, "1", "0")
End Sub

'--- المفتاح محفوظ لكل مزوّد على حدة، ومموّه (تمويه لا تشفير) --------
Public Function CfgApiKey(Optional ByVal provider As String = "") As String
    If Len(provider) = 0 Then provider = CfgProvider()
    CfgApiKey = Deobfuscate(CfgGet("Key_" & provider, ""))
End Function

Public Sub CfgSetApiKey(ByVal provider As String, ByVal keyValue As String)
    CfgSet "Key_" & provider, Obfuscate(Trim$(keyValue))
End Sub

'--- النموذج محفوظ لكل مزوّد على حدة --------------------------------
Public Function CfgModel(Optional ByVal provider As String = "") As String
    If Len(provider) = 0 Then provider = CfgProvider()
    CfgModel = CfgGet("Model_" & provider, DefaultModel(provider))
End Function

Public Sub CfgSetModel(ByVal provider As String, ByVal modelId As String)
    CfgSet "Model_" & provider, Trim$(modelId)
End Sub

Public Function DefaultModel(ByVal provider As String) As String
    ' الافتراضي دائماً نموذج مجاني — ما في أي كلفة على المستخدم.
    ' معرّفات النماذج بتتغير مع الوقت، فزر "تحديث قائمة النماذج"
    ' في الإعدادات بيجيب القائمة الحيّة من المزوّد.
    Select Case LCase$(provider)
        Case "omniroute":  DefaultModel = OMNI_MODEL
        Case "nvidia":     DefaultModel = "nvidia/nemotron-3-super-120b-a12b"
        Case "openrouter": DefaultModel = "nvidia/nemotron-3-ultra-550b-a55b:free"
        Case "groq":       DefaultModel = "llama-3.3-70b-versatile"
        Case "anthropic":  DefaultModel = "claude-haiku-4-5"
        Case "custom":     DefaultModel = PRESET_MODEL
        Case Else:         DefaultModel = ""
    End Select
End Function

'====================================================================
' إعدادات السلوك
'====================================================================
Public Function CfgAllowMacros() As Boolean
    CfgAllowMacros = (CfgGet("AllowMacros", "1") = "1")
End Function

Public Sub CfgSetAllowMacros(ByVal b As Boolean)
    CfgSet "AllowMacros", IIf(b, "1", "0")
End Sub

Public Function CfgAutoBackup() As Boolean
    CfgAutoBackup = (CfgGet("AutoBackup", "1") = "1")
End Function

Public Sub CfgSetAutoBackup(ByVal b As Boolean)
    CfgSet "AutoBackup", IIf(b, "1", "0")
End Sub

Public Function CfgMaxSteps() As Long
    Dim n As Long
    n = CLng(Val(CfgGet("MaxSteps", "14")))
    If n < 3 Then n = 3
    If n > 40 Then n = 40
    CfgMaxSteps = n
End Function

Public Function CfgTemperature() As Double
    Dim d As Double
    d = Val(Replace(CfgGet("Temperature", "0.2"), ",", "."))
    If d < 0 Then d = 0
    If d > 1 Then d = 1
    CfgTemperature = d
End Function

Public Function IsConfigured() As Boolean
    ' OmniRoute بتشتغل بدون مفتاح
    If CfgProvider() = "omniroute" Then
        IsConfigured = (Len(CfgModel()) > 0)
    Else
        IsConfigured = (Len(CfgApiKey()) > 5) And (Len(CfgModel()) > 0)
    End If
End Function

'--- هل نستخدم استدعاء الأدوات الأصلي مع المزوّد المخصّص؟ ------------
Public Function CfgNativeTools() As Boolean
    CfgNativeTools = (CfgGet("NativeTools", "1") = "1")
End Function

Public Sub CfgSetNativeTools(ByVal b As Boolean)
    CfgSet "NativeTools", IIf(b, "1", "0")
End Sub

'====================================================================
' تمويه بسيط للمفتاح داخل السجل
' ملاحظة: هذا تمويه لمنع القراءة العابرة، وليس تشفيراً أمنياً.
'====================================================================
Private Const OBF_SEED As String = "MaliAI-2026-Excel"

Private Function Obfuscate(ByVal s As String) As String
    If Len(s) = 0 Then Obfuscate = "": Exit Function
    Dim i As Long, c As Long, k As Long
    Dim sb As String
    For i = 1 To Len(s)
        k = AscW(Mid$(OBF_SEED, ((i - 1) Mod Len(OBF_SEED)) + 1, 1))
        c = AscW(Mid$(s, i, 1)) Xor k
        sb = sb & Right$("000" & Hex$(c), 4)
    Next i
    Obfuscate = sb
End Function

Private Function Deobfuscate(ByVal s As String) As String
    On Error GoTo Fail
    If Len(s) = 0 Then Deobfuscate = "": Exit Function
    If Len(s) Mod 4 <> 0 Then Deobfuscate = s: Exit Function   ' مفتاح قديم غير مموّه
    Dim i As Long, idx As Long, c As Long, k As Long
    Dim sb As String
    idx = 0
    For i = 1 To Len(s) Step 4
        idx = idx + 1
        k = AscW(Mid$(OBF_SEED, ((idx - 1) Mod Len(OBF_SEED)) + 1, 1))
        c = CLng("&H" & Mid$(s, i, 4)) Xor k
        sb = sb & ChrW$(c)
    Next i
    Deobfuscate = sb
    Exit Function
Fail:
    Deobfuscate = ""
End Function
