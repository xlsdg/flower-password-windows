VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flower Password"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":43B2
   ScaleHeight     =   3090
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox chkAddUserCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "On"
      Height          =   255
      Left            =   3380
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "附加扰码临时开关"
      Top             =   480
      Value           =   1  'Checked
      Width           =   600
   End
   Begin VB.PictureBox picKey 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   240
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   300
      ScaleWidth      =   1725
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "用于区别不同用途密码的简短代号，如淘宝账号可用“taobao”或“tb”等"
      Top             =   1463
      Width           =   1725
   End
   Begin VB.PictureBox picSetting 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   4320
      Picture         =   "FrmMain.frx":64E1
      ScaleHeight     =   2505
      ScaleWidth      =   4050
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4050
      Begin VB.PictureBox picUserCode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   340
         MousePointer    =   3  'I-Beam
         ScaleHeight     =   300
         ScaleWidth      =   1725
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "自动在区分代号中添加附加扰码"
         Top             =   1963
         Width           =   1725
      End
      Begin VB.ComboBox comUserCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   3570
      End
      Begin VB.CheckBox chkDomainSuffix 
         BackColor       =   &H00FFFFFF&
         Caption         =   "域名后缀"
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         ToolTipText     =   "将包含后缀的域名填入区分代号"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optSuffix 
         BackColor       =   &H00F6F6F6&
         Caption         =   "后缀"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         ToolTipText     =   "区分代号+附加扰码"
         Top             =   1500
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optPrefix 
         BackColor       =   &H00F6F6F6&
         Caption         =   "前缀"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "附加扰码+区分代号"
         Top             =   1500
         Width           =   735
      End
      Begin VB.CheckBox chkAutoCopy 
         BackColor       =   &H00FFFFFF&
         Caption         =   "自动将生成的密码复制到剪贴板"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkAutoAddUserCode 
         BackColor       =   &H00F6F6F6&
         Caption         =   "自动添加附加扰码："
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAutoUseDomain 
         BackColor       =   &H00FAFAFA&
         Caption         =   "自动将域名填入区分代号"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Image iSave 
         Height          =   195
         Left            =   3480
         Picture         =   "FrmMain.frx":84C8
         Top             =   120
         Width           =   195
      End
      Begin VB.Image imgSave 
         Height          =   195
         Left            =   3720
         MouseIcon       =   "FrmMain.frx":854E
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":86A4
         ToolTipText     =   "保存"
         Top             =   120
         Width           =   195
      End
   End
   Begin VB.Timer tmrZip 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picPassword 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   240
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   300
      ScaleWidth      =   1725
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "可选择一个简单易记的密码，用于生成其他高强度密码"
      Top             =   920
      Width           =   1725
   End
   Begin VB.ComboBox comKey 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   140
      TabIndex        =   1
      ToolTipText     =   "用于区别不同用途密码的简短代号，如淘宝账号可用“taobao”或“tb”等"
      Top             =   1420
      Width           =   3760
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   140
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "可选择一个简单易记的密码，用于生成其他高强度密码"
      Top             =   877
      Width           =   3760
   End
   Begin VB.Label lblStrength 
      Alignment       =   2  'Center
      BackColor       =   &H00EBEBEB&
      Caption         =   "Not rated"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   173
      Left            =   135
      TabIndex        =   15
      ToolTipText     =   "密码强度提示"
      Top             =   1257
      Width           =   3765
   End
   Begin VB.Image imgCopy 
      Height          =   360
      Left            =   1905
      Picture         =   "FrmMain.frx":87F5
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image iClose 
      Height          =   195
      Left            =   4080
      Picture         =   "FrmMain.frx":88BD
      Top             =   120
      Width           =   195
   End
   Begin VB.Image iSetting 
      Height          =   315
      Left            =   3600
      Picture         =   "FrmMain.frx":8943
      Top             =   2640
      Width           =   315
   End
   Begin VB.Image iInfo 
      Height          =   315
      Left            =   3240
      Picture         =   "FrmMain.frx":8A5A
      Top             =   2640
      Width           =   330
   End
   Begin VB.Image iHome 
      Height          =   315
      Left            =   2880
      Picture         =   "FrmMain.frx":8B20
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label lblCode16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   120
      MouseIcon       =   "FrmMain.frx":8BE7
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "点击复制十六位码"
      Top             =   2130
      Width           =   60
   End
   Begin VB.Image imgInfo 
      Height          =   315
      Left            =   3210
      MouseIcon       =   "FrmMain.frx":8D39
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":8E8F
      ToolTipText     =   "帮助"
      Top             =   2070
      Width           =   330
   End
   Begin VB.Image imgHome 
      Height          =   315
      Left            =   2842
      MouseIcon       =   "FrmMain.frx":9040
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":9196
      ToolTipText     =   "官网"
      Top             =   2070
      Width           =   315
   End
   Begin VB.Image imgSetting 
      Height          =   315
      Left            =   3610
      MouseIcon       =   "FrmMain.frx":9412
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":9568
      ToolTipText     =   "设置"
      Top             =   2070
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   195
      Left            =   3720
      MouseIcon       =   "FrmMain.frx":981F
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":9975
      ToolTipText     =   "关闭"
      Top             =   120
      Width           =   195
   End
   Begin VB.Menu munFlowerPassword 
      Caption         =   "花密(&F)"
      Visible         =   0   'False
      Begin VB.Menu munShow 
         Caption         =   "显示(&S)"
      End
      Begin VB.Menu munLine 
         Caption         =   "-"
      End
      Begin VB.Menu munExit 
         Caption         =   "退出(&E)"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FrmHide(ByVal isPostKey As Boolean, ByVal isKeyPress As Boolean)
    UnSetWinOnTop Me.hwnd
    Call UnSetMouseHook
    Me.Visible = False

    If isPostKey Then
        If Len(lblCode16.Caption) > 0 Then
            PoseCodeToClipboard lblCode16.Caption

            If isKeyPress Then
                PostCode lblCode16.Caption, Password_Hwnd
            Else
                PostCode lblCode16.Caption, 0

            End If

            AddStrToCombox comKey, GetKeyFromStr(comKey.Text)

        End If

    Else

        If Len(lblCode16.Caption) > 0 Then
            PoseCodeToClipboard lblCode16.Caption

        End If

    End If

    txtPassword.Text = ""

End Sub

Public Sub FrmShow()

    If SetMouseHook = 0 Then
        MsgBox "花密加载鼠标挂钩模块失败！", vbCritical + vbOKOnly + vbSystemModal

    End If

    If SetWinOnTop(Me.hwnd) = 0 Then
        MsgBox "花密窗体置顶失败！", vbCritical + vbOKOnly + vbSystemModal

    End If

    Me.Visible = True
    txtPassword.SetFocus

End Sub

Public Sub MidScreenShow()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Call FrmShow

End Sub

Public Sub ShowFrmByHotKey()

    If Not Me.Visible Then
        If isSetting Then
            comUserCode.Clear
            Call LoadSetting
            Call ShowResult
            Call HideSetting

        End If

        Dim x As Long, Y As Long

        getLocation x, Y

        If isAutoUseDomain Then
            SetUrlAsKey Password_Hwnd

        End If

        Me.Move x * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
        Call FrmShow

    End If

End Sub

Private Sub HideSetting()
    chkAddUserCode.Visible = True
    picPassword.Top = 920

    If Len(txtPassword.Text) = 0 Then
        picPassword.Visible = True

    End If

    picKey.Top = 1463

    If Len(comKey.Text) = 0 Then
        picKey.Visible = True

    End If

    picSetting.Visible = False
    'txtPassword.SetFocus
    isSetting = False

End Sub

Private Sub ShowPasswordStrength()

    If Len(txtPassword.Text) = 0 Then
        lblStrength.Caption = "Not rated"
        lblStrength.BackColor = &HEBEBEB
    Else

        Dim strHelp As String

        Select Case check_password_level(txtPassword.Text, strHelp)

            Case PASSWORD_INSECURE
                lblStrength.Caption = "Weak"
                lblStrength.BackColor = &H4545FF

            Case PASSWORD_WARNING
                lblStrength.Caption = "Medium"
                lblStrength.BackColor = &H5ED3FF

            Case PASSWORD_ADVICE
                lblStrength.Caption = "Strong"
                lblStrength.BackColor = &HCE9A02

            Case PASSWORD_ACHIEVEMENT
                lblStrength.Caption = "Best"
                lblStrength.BackColor = &H1CBB3A

        End Select

        lblStrength.ToolTipText = strHelp

    End If

End Sub

Private Sub ShowResult()
    KeyBox_AutoComplete comKey

    If Len(txtPassword.Text) > 0 And Len(comKey.Text) > 0 Then
        'If isAutoAddUserCode Then
        'If isPrefix Then
        'lblCode16.Caption = getFlowerPassword(txtPassword.Text, comUserCode.Text + comKey.Text, 16)
        'Else
        'lblCode16.Caption = getFlowerPassword(txtPassword.Text, comKey.Text + comUserCode.Text, 16)
        'End If
        'Else
        lblCode16.Caption = getFlowerPassword(txtPassword.Text, comKey.Text, 16)
        'End If
    Else
        lblCode16.Caption = ""

    End If

End Sub

Private Sub ShowSetting()
    chkAddUserCode.Visible = False
    picPassword.Top = 2640: picPassword.Visible = False
    picKey.Top = 2640: picKey.Visible = False
    picSetting.Left = 0: picSetting.Top = 0
    picSetting.Visible = True

    If comUserCode.Enabled Then
        comUserCode.SetFocus

    End If

    isSetting = True

End Sub

Private Sub chkAddUserCode_Click()

    If chkAddUserCode.value = Checked Then
        isAutoAddUserCode = True
        chkAutoAddUserCode.value = Checked
        chkAddUserCode.ForeColor = vbGreen
        chkAddUserCode.Caption = "On"
    Else
        isAutoAddUserCode = False
        chkAutoAddUserCode.value = Unchecked
        chkAddUserCode.ForeColor = vbRed
        chkAddUserCode.Caption = "Off"

    End If

    Call ShowResult

    'Call SaveSetting
End Sub

Private Sub chkAutoAddUserCode_Click()

    If chkAutoAddUserCode.value = Checked Then
        optPrefix.Enabled = True
        optSuffix.Enabled = True
        comUserCode.Enabled = True
        picUserCode.Enabled = True
        chkAddUserCode.value = Checked
    Else
        optPrefix.Enabled = False
        optSuffix.Enabled = False
        comUserCode.Enabled = False
        picUserCode.Enabled = False
        chkAddUserCode.value = Unchecked

    End If

End Sub

Private Sub chkAutoUseDomain_Click()

    If chkAutoUseDomain.value = Checked Then
        chkDomainSuffix.Enabled = True
    Else
        chkDomainSuffix.Enabled = False

    End If

End Sub

Private Sub comKey_Change()

    If Len(comKey.Text) > 0 Then
        picKey.Visible = False
    Else
        picKey.Visible = True

    End If

    Call ShowResult

End Sub

Private Sub comKey_Click()

    If Len(comKey.Text) > 0 Then
        picKey.Visible = False
    Else
        picKey.Visible = True

    End If

    Call ShowResult

End Sub

Private Sub comKey_GotFocus()

    'picKey.Visible = False
    'If Len(comKey.Text) > 0 Then
    'comKey.ForeColor = &H80000008
    'comKey.Text = AddCodeToKey(comKey.Text, txtUserCode.Text)
    'comKey.SelStart = 0
    'comKey.SelLength = Len(comKey.Text)
    'End If
End Sub

Private Sub comKey_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyBox_KeyDown comKey, KeyCode, Shift

    If KeyCode = vbKeyEscape Then
        FrmHide False, True
    ElseIf KeyCode = vbKeyReturn Then

        If Len(lblCode16.Caption) > 0 Then
            FrmHide True, True
        Else
            FrmHide False, True

        End If

    ElseIf Len(comKey.Text) > 0 Then
        picKey.Visible = False

    End If

End Sub

Private Sub comKey_LostFocus()

    If Len(comKey.Text) > 0 Then
        'comKey.ForeColor = &H808080
        'comKey.Text = AddCodeToKey(comKey.Text, txtUserCode.Text)
        picKey.Visible = False
    Else
        picKey.Visible = True

    End If

End Sub

Private Sub comUserCode_Change()

    If Len(comUserCode.Text) > 0 Then
        picUserCode.Visible = False
    Else
        picUserCode.Visible = True

    End If

    If Not isLoading Then cbBox_Change comUserCode

End Sub

Private Sub comUserCode_Click()

    If Len(comUserCode.Text) > 0 Then
        picUserCode.Visible = False
    Else
        picUserCode.Visible = True

    End If

End Sub

Private Sub comUserCode_GotFocus()

    'picUser.Visible = False
    'If Len(comUserCode.Text) > 0 Then
    'comUserCode.SelStart = 0
    'comUserCode.SelLength = Len(comUserCode.Text)
    'End If
End Sub

Private Sub comUserCode_KeyDown(KeyCode As Integer, Shift As Integer)
    cbBox_KeyDown comUserCode, KeyCode, Shift

    If KeyCode = vbKeyEscape Then
        FrmHide False, True
    ElseIf KeyCode = vbKeyReturn Then
        Call imgSave_Click
    ElseIf Len(comUserCode.Text) > 0 Then
        picUserCode.Visible = False

    End If

End Sub

Private Sub comUserCode_LostFocus()

    If Len(comUserCode.Text) > 0 Then
        picUserCode.Visible = False
    Else
        picUserCode.Visible = True

    End If

    cbBox_LostFocus comUserCode

End Sub

Private Sub Form_Initialize()

    If Not isRunInIDEMode Then
        If App.PrevInstance Then
            MsgBox "花密已经运行，请关闭后台程序再重试一次！", vbCritical + vbOKOnly + vbSystemModal
            End
        Else
            App.TaskVisible = False
            Call InitCommonControlsVB

            If ProtectTextBox(txtPassword.hwnd) = 0 Then
                MsgBox "花密记忆密码输入框保护失败！", vbCritical + vbOKOnly + vbSystemModal

            End If

        End If

    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
        FrmHide False, True

    End If

End Sub

Private Sub Form_Load()
    picPassword.Print "请输入记忆密码": picKey.Print "请输入区分代号": picUserCode.Print "请输入附加扰码"
    SetWinStyle Me.hwnd: Me.Width = 4060: Me.Height = 2513
    iClose.Left = imgClose.Left: iClose.Top = imgClose.Top: iHome.Left = imgHome.Left: iHome.Top = imgHome.Top: iInfo.Left = imgInfo.Left: iInfo.Top = imgInfo.Top: iSetting.Left = imgSetting.Left: iSetting.Top = imgSetting.Top: iSave.Left = imgSave.Left: iSave.Top = imgSave.Top

    If SetHotKey(Me.hwnd) = 0 Then
        MsgBox "花密快捷键（WIN + S）与系统其它程序冲突！", vbCritical + vbOKOnly + vbSystemModal
        End

    End If

    If AddToTray(FrmMain, App.Title) = 0 Then
        MsgBox "花密加载托盘模块失败！", vbCritical + vbOKOnly + vbSystemModal
        End
    Else
        SetTrayMsgbox "按快捷键 WIN + S 呼唤花密，按 Enter 键或 Esc 键关闭输入窗口。", NIIF_GUID, "欢迎您使用花密Windows版", Me.Icon

    End If

    If SetFormRgn = 0 Then
        MsgBox "花密窗体特效设置失败！", vbCritical + vbOKOnly + vbSystemModal
        End

    End If

    Call LoadSetting

    If isRunInIDEMode Then
        tmrZip.Enabled = False
    Else
        Call ZipMemory

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SetWinMove Me.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    iClose.Visible = True: iHome.Visible = True: iInfo.Visible = True: iSetting.Visible = True: imgCopy.Visible = False

    If Not Me.Visible Then
        MouseOnTray Button, Shift, x, Y

    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    txtPassword.SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call UnSetFormRgn
    UnSetHotKey Me.hwnd

End Sub

Private Sub Form_Terminate()
    Call RemoveFromTray

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not isRunInIDEMode Then UnProtectTextBox txtPassword.hwnd

End Sub

Private Sub iClose_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             Y As Single)
    iClose.Visible = False

End Sub

Private Sub iHome_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    iHome.Visible = False

End Sub

Private Sub iInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    iInfo.Visible = False

End Sub

Private Sub imgClose_Click()
    iClose.Visible = True
    FrmHide False, False

End Sub

Private Sub imgHome_Click()
    OpenWebsite "http://flowerpassword.com"
    iHome.Visible = True

End Sub

Private Sub imgInfo_Click()
    OpenWebsite "http://flowerpassword.com/guide"
    iInfo.Visible = True

End Sub

Private Sub imgSave_Click()
    iSave.Visible = True
    Call SaveSetting
    Call ShowResult
    Call HideSetting

End Sub

Private Sub imgSetting_Click()
    iSetting.Visible = True
    Call ShowSetting

End Sub

Private Sub iSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    iSave.Visible = False

End Sub

Private Sub iSetting_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)
    iSetting.Visible = False

End Sub

Private Sub lblCode16_Click()

    If Len(lblCode16.Caption) > 0 Then
        Clipboard.Clear
        Clipboard.SetText lblCode16.Caption
        imgCopy.Visible = True

    End If

End Sub

Private Sub munExit_Click()
    Unload Me

End Sub

Private Sub munShow_Click()
    Call MidScreenShow

End Sub

Private Sub picKey_Click()
    'picKey.Visible = False
    comKey.SetFocus

End Sub

Private Sub picPassword_Click()
    'picPassword.Visible = False
    txtPassword.SetFocus

End Sub

Private Sub picSetting_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
        FrmHide False, True

    End If

End Sub

Private Sub picSetting_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 Y As Single)
    SetWinMove Me.hwnd

End Sub

Private Sub picSetting_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 Y As Single)
    iSave.Visible = True

End Sub

Private Sub picSetting_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)

    If comUserCode.Enabled Then comUserCode.SetFocus

End Sub

Private Sub picUserCode_Click()
    comUserCode.SetFocus

End Sub

Private Sub tmrZip_Timer()

    Static I As Integer

    I = I + 1

    If I > 60 Then
        I = 0
        Call ZipMemory

    End If

End Sub

Private Sub txtPassword_Change()
    Call ShowPasswordStrength

    If Len(txtPassword.Text) > 0 Then
        picPassword.Visible = False
    Else
        picPassword.Visible = True

    End If

    Call ShowResult

End Sub

Private Sub txtPassword_GotFocus()

    'picPassword.Visible = False
    If Len(txtPassword.Text) > 0 Then
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)

    End If

End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        FrmHide False, True
    ElseIf KeyCode = vbKeyReturn Then

        If Len(lblCode16.Caption) > 0 And Len(comKey.Text) > 0 Then
            FrmHide True, True
        Else
            comKey.SetFocus

        End If

    ElseIf Len(txtPassword.Text) > 0 Then
        picPassword.Visible = False

    End If

End Sub

Private Sub txtPassword_LostFocus()

    If Len(txtPassword.Text) > 0 Then
        picPassword.Visible = False
    Else
        picPassword.Visible = True

    End If

End Sub

