Attribute VB_Name = "modUI"
'====================================================================
' modUI  -  نقاط الدخول، الشريط (Ribbon)، والقائمة الاحتياطية
'====================================================================
Option Explicit

Public gChat As frmChat
Public gRibbonLoaded As Boolean

Private Const BAR_NAME As String = "MaliAI_Bar"

'====================================================================
' فتح نافذة المساعد
'====================================================================
Public Sub ShowAssistant()
    On Error GoTo Fail

    If Not IsConfigured() Then
        If MsgBox("لسا ما ضبطت مفتاح الذكاء الاصطناعي." & vbCrLf & vbCrLf & _
                  "بدك تفتح الإعدادات هلأ؟", vbYesNo + vbInformation + vbMsgBoxRight, _
                  APP_TITLE) = vbYes Then
            ShowSettings
        End If
        If Not IsConfigured() Then Exit Sub
    End If

    If gChat Is Nothing Then Set gChat = New frmChat
    gChat.Show vbModeless
    On Error Resume Next
    gChat.FocusInput
    Exit Sub

Fail:
    ' النموذج انسكّر من المستخدم — أنشئه من جديد
    Set gChat = New frmChat
    gChat.Show vbModeless
End Sub

Public Sub ShowSettings()
    Dim f As frmSettings
    Set f = New frmSettings
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub

'====================================================================
' أوامر سريعة
'====================================================================
Public Sub QuickExplainCell()
    If Not CheckReady() Then Exit Sub
    Dim sel As Object
    Set sel = Application.Selection
    If TypeName(sel) <> "Range" Then
        MsgBox "حدّد خلية أو نطاق أولاً.", vbInformation + vbMsgBoxRight, APP_TITLE
        Exit Sub
    End If
    ShowAssistant
    On Error Resume Next
    gChat.AskNow "اشرحلي بالتفصيل شو بتعمل الخلايا المحددة (" & _
                 sel.Address(False, False) & ") وشو منطق المعادلة، وإذا في فيها خطأ أو خطر قوللي."
End Sub

Public Sub QuickAuditSheet()
    If Not CheckReady() Then Exit Sub
    ShowAssistant
    On Error Resume Next
    gChat.AskNow "دقّق الورقة النشطة كمراجع مالي: دوّر على أخطاء المعادلات (#REF!, #VALUE!, #DIV/0!)، " & _
                 "أرقام مخزّنة كنص، صفوف أو أعمدة فاضية جوّا البيانات، مجاميع ما بتطابق تفاصيلها، " & _
                 "تواريخ غير موحّدة، وخلايا مدمجة بتخرب الحسابات. " & _
                 "اقرأ البيانات أول شي، وتحقق من المجاميع بأداة precise_sum، " & _
                 "وبالآخر اعطيني تقرير مختصر بالمشاكل مرتبة حسب الخطورة مع مكان كل مشكلة."
End Sub

Public Sub QuickBuildFormula()
    If Not CheckReady() Then Exit Sub
    ShowAssistant
    On Error Resume Next
    gChat.PrefillInput "بدي معادلة تعمل: "
End Sub

Public Sub DoUndo()
    Dim msg As String
    msg = UndoLast()
    MsgBox msg, vbInformation + vbMsgBoxRight, APP_TITLE & " — تراجع"
End Sub

Public Sub DoBackup()
    Dim m As String
    m = BackupActiveBook()
    If Len(m) = 0 Then m = "النسخ الاحتياطي معطّل من الإعدادات."
    MsgBox m, vbInformation + vbMsgBoxRight, APP_TITLE
End Sub

Private Function CheckReady() As Boolean
    If ActiveWorkbook Is Nothing Then
        MsgBox "افتح ملف إكسل أولاً.", vbInformation + vbMsgBoxRight, APP_TITLE
        CheckReady = False
        Exit Function
    End If
    CheckReady = True
End Function

'====================================================================
' المساعدة
'====================================================================
Public Sub ShowHelp()
    Dim s As String
    s = APP_TITLE & "  —  الإصدار " & APP_VERSION & vbCrLf & String(46, "=") & vbCrLf & vbCrLf
    s = s & "شو بيعمل؟" & vbCrLf
    s = s & "بتحكيله بالعربي شو بدك، وهو بيشتغل على الشيت مباشرة:" & vbCrLf
    s = s & "  • بيكتب المعادلات وبيشرحها" & vbCrLf
    s = s & "  • بينظّف ويرتّب البيانات ويشيل المكرر" & vbCrLf
    s = s & "  • بيعمل مجاميع وتقارير ومطابقات وأعمار ديون" & vbCrLf
    s = s & "  • بيعمل رسوم بيانية" & vbCrLf
    s = s & "  • بيكتب وبيشغّل ماكرو بعد موافقتك" & vbCrLf & vbCrLf
    s = s & "أمثلة تكتبها:" & vbCrLf
    s = s & "  - ""اعملي عمود ضريبة ١٦٪ جنب عمود المبلغ""" & vbCrLf
    s = s & "  - ""طابق أرصدة عمود D مع عمود H وقلي وين الفرق""" & vbCrLf
    s = s & "  - ""رتب الجدول حسب التاريخ وشيل المكرر""" & vbCrLf
    s = s & "  - ""اعملي تحليل أعمار ديون لكل عميل: أقل من ٣٠، ٦٠، ٩٠، أكثر""" & vbCrLf
    s = s & "  - ""ليش هاي المعادلة بترجع #VALUE؟""" & vbCrLf & vbCrLf
    s = s & "الأمان:" & vbCrLf
    s = s & "  • أي تعديل بيقدر يترجع بزر ""تراجع"" (آخر ٢٠ خطوة)" & vbCrLf
    s = s & "  • قبل أي ماكرو بتنحفظ نسخة احتياطية من الملف" & vbCrLf
    s = s & "  • الماكرو ما بيشتغل إلا بعد ما تشوف الكود وتوافق" & vbCrLf
    s = s & "  • المفتاح محفوظ على جهازك بس، وما بينبعت لأي جهة غير المزوّد" & vbCrLf & vbCrLf
    s = s & "المزوّد الحالي: " & ProviderTitle(CfgProvider()) & vbCrLf
    s = s & "النموذج الحالي: " & CfgModel() & vbCrLf & vbCrLf
    s = s & "اختصار سريع:  Ctrl + Shift + A"

    MsgBox s, vbInformation + vbMsgBoxRight, APP_TITLE & " — مساعدة"
End Sub

'====================================================================
' استدعاءات الشريط (Ribbon)
'====================================================================
' ملاحظة: الوسائط معرّفة As Object عمداً حتى ما نعتمد على مرجع
' مكتبة Office، فالإضافة بتشتغل على أي نسخة إكسل بدون ضبط مراجع.

Public Sub RibbonOnLoad(ByVal ribbon As Object)
    gRibbonLoaded = True
End Sub

Public Sub Rx_Open(ByVal control As Object)
    ShowAssistant
End Sub

Public Sub Rx_Explain(ByVal control As Object)
    QuickExplainCell
End Sub

Public Sub Rx_Audit(ByVal control As Object)
    QuickAuditSheet
End Sub

Public Sub Rx_Formula(ByVal control As Object)
    QuickBuildFormula
End Sub

Public Sub Rx_Undo(ByVal control As Object)
    DoUndo
End Sub

Public Sub Rx_Backup(ByVal control As Object)
    DoBackup
End Sub

Public Sub Rx_Settings(ByVal control As Object)
    ShowSettings
End Sub

Public Sub Rx_Help(ByVal control As Object)
    ShowHelp
End Sub

'====================================================================
' قائمة احتياطية (تظهر في تبويب Add-ins) إذا ما اشتغل الشريط
'====================================================================
Public Sub BuildFallbackMenu()
    On Error Resume Next
    RemoveFallbackMenu

    Dim bar As Object
    Set bar = Application.CommandBars.Add(Name:=BAR_NAME, Position:=1, Temporary:=True)

    AddBtn bar, APP_TITLE, "ShowAssistant"
    AddBtn bar, "شرح الخلية", "QuickExplainCell"
    AddBtn bar, "تدقيق الورقة", "QuickAuditSheet"
    AddBtn bar, "تراجع", "DoUndo"
    AddBtn bar, "الإعدادات", "ShowSettings"
    AddBtn bar, "مساعدة", "ShowHelp"

    bar.Visible = True
End Sub

Private Sub AddBtn(ByVal bar As Object, ByVal cap As String, ByVal act As String)
    On Error Resume Next
    Dim b As Object
    Set b = bar.Controls.Add(Type:=1, Temporary:=True)
    b.Caption = cap
    b.OnAction = act
    b.Style = 2                     ' msoButtonCaption
End Sub

Public Sub RemoveFallbackMenu()
    On Error Resume Next
    Application.CommandBars(BAR_NAME).Delete
End Sub

'--- تُستدعى بعد ثانية من الفتح: إذا الشريط ما حمّل، نركّب القائمة ---
Public Sub CheckRibbonFallback()
    If Not gRibbonLoaded Then BuildFallbackMenu
End Sub
