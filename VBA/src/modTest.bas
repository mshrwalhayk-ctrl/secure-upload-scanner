Attribute VB_Name = "modTest"
'====================================================================
' modTest  -  فحص ذاتي للإضافة (بدون إنترنت وبدون مفتاح)
'             شغّله من: Alt+F8  ->  MaliAI_SelfTest
'====================================================================
Option Explicit

Private mPass As Long
Private mFail As Long
Private mLog As String

Public Sub MaliAI_SelfTest()
    mPass = 0: mFail = 0: mLog = ""

    TestJson
    TestMathEngine
    TestConfig

    Dim head As String
    head = APP_TITLE & " — نتيجة الفحص الذاتي" & vbCrLf & String(40, "=") & vbCrLf
    head = head & "ناجح: " & mPass & "    فاشل: " & mFail & vbCrLf & vbCrLf

    MsgBox head & mLog, IIf(mFail = 0, vbInformation, vbExclamation) + vbMsgBoxRight, APP_TITLE
End Sub

'--------------------------------------------------------------------
Private Sub Chk(ByVal name As String, ByVal actual As String, ByVal expected As String)
    If actual = expected Then
        mPass = mPass + 1
        mLog = mLog & "✔ " & name & vbCrLf
    Else
        mFail = mFail + 1
        mLog = mLog & "✖ " & name & vbCrLf & _
               "     المتوقع: " & expected & vbCrLf & _
               "     الفعلي : " & actual & vbCrLf
    End If
End Sub

'--------------------------------------------------------------------
Private Sub TestJson()
    On Error GoTo Fail

    Dim j As Variant
    Set j = JsonTryParse("{""a"":1,""b"":[10,20,30],""c"":{""d"":""نص عربي""},""e"":true,""f"":null}")

    If j Is Nothing Then
        mFail = mFail + 1
        mLog = mLog & "✖ JSON: فشل التحليل من الأساس" & vbCrLf
        Exit Sub
    End If

    Chk "JSON: قراءة رقم", JGetStr(j, "a"), "1"
    Chk "JSON: قراءة عنصر من مصفوفة", JGetStr(j, "b/1"), "20"
    Chk "JSON: قراءة نص عربي متداخل", JGetStr(j, "c/d"), "نص عربي"
    Chk "JSON: قيمة منطقية", LCase$(JGetStr(j, "e")), LCase$(CStr(True))
    Chk "JSON: مفتاح غير موجود", JGetStr(j, "zzz"), ""

    ' تهريب العربي إلى \u
    Chk "JSON: تهريب حرف عربي", JsonEscape("م"), "\u0645"
    Chk "JSON: تهريب علامة تنصيص", JsonEscape("a""b"), "a\""b"
    Chk "JSON: تهريب سطر جديد", JsonEscape("a" & vbLf & "b"), "a\nb"

    ' دورة كاملة: بناء ثم تحليل
    Dim built As String
    built = "{""msg"":" & JsonStr("مرحبا " & vbLf & """اقتباس""") & "}"
    Dim k As Variant
    Set k = JsonTryParse(built)
    If k Is Nothing Then
        mFail = mFail + 1
        mLog = mLog & "✖ JSON: فشل في الدورة الكاملة" & vbCrLf
    Else
        Chk "JSON: دورة بناء/تحليل كاملة", JGetStr(k, "msg"), "مرحبا " & vbLf & """اقتباس"""
    End If

    ' استخراج نص من رد على شكل OpenAI
    Dim resp As String
    resp = "{""choices"":[{""message"":{""role"":""assistant"",""content"":""تمام""}}]}"
    Dim r As Variant
    Set r = JsonTryParse(resp)
    Chk "JSON: استخراج رد النموذج", JGetStr(r, "choices/0/message/content"), "تمام"
    Exit Sub

Fail:
    mFail = mFail + 1
    mLog = mLog & "✖ JSON: خطأ غير متوقع — " & Err.Description & vbCrLf
End Sub

'--------------------------------------------------------------------
Private Sub TestMathEngine()
    On Error GoTo Fail
    Dim okFlag As Boolean

    Chk "حساب: ضرب بسيط", MathEval("1250*1.16", okFlag), "1450"
    Chk "حساب: أقواس وأولويات", MathEval("(2+3)*4-10/5", okFlag), "18"
    Chk "حساب: دالة إكسل ROUND", MathEval("ROUND(2.34567,2)", okFlag), "2.35"
    Chk "حساب: أرقام عربية ٢٥+٧٥", MathEval("٢٥+٧٥", okFlag), "100"
    Chk "حساب: علامة الضرب ×", MathEval("12×12", okFlag), "144"
    Chk "حساب: معادلة مع = في البداية", MathEval("=50/4", okFlag), "12.5"

    ' القسمة على صفر لازم تُمسك كخطأ لا كنتيجة
    Dim divRes As String
    divRes = MathEval("1/0", okFlag)
    If okFlag Then
        mFail = mFail + 1
        mLog = mLog & "✖ حساب: القسمة على صفر ما اتمسكت" & vbCrLf
    Else
        mPass = mPass + 1
        mLog = mLog & "✔ حساب: القسمة على صفر اتمسكت كخطأ" & vbCrLf
    End If

    ' الدقة المالية: 0.1+0.2 لازم تطلع 0.3 بالضبط
    Chk "حساب: دقة مالية 0.1+0.2", NumToText(CDec("0.1") + CDec("0.2")), "0.3"
    Exit Sub

Fail:
    mFail = mFail + 1
    mLog = mLog & "✖ حساب: خطأ غير متوقع — " & Err.Description & vbCrLf
End Sub

'--------------------------------------------------------------------
Private Sub TestConfig()
    On Error GoTo Fail

    Dim saved As String
    saved = CfgApiKey("groq")

    CfgSetApiKey "groq", "sk-test-مفتاح-123"
    Chk "الإعدادات: حفظ واسترجاع المفتاح", CfgApiKey("groq"), "sk-test-مفتاح-123"

    ' التأكد إنه المفتاح مش مخزّن كنص صريح في السجل
    Dim rawStored As String
    rawStored = CfgGet("Key_groq", "")
    If InStr(rawStored, "sk-test") > 0 Then
        mFail = mFail + 1
        mLog = mLog & "✖ الإعدادات: المفتاح مخزّن كنص صريح" & vbCrLf
    Else
        mPass = mPass + 1
        mLog = mLog & "✔ الإعدادات: المفتاح مموّه في السجل" & vbCrLf
    End If

    CfgSetApiKey "groq", saved       ' إرجاع القيمة الأصلية
    Exit Sub

Fail:
    mFail = mFail + 1
    mLog = mLog & "✖ الإعدادات: خطأ غير متوقع — " & Err.Description & vbCrLf
End Sub
