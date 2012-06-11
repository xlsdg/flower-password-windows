VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Flower Password"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
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
   ScaleHeight     =   2520
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox chkAddUserCode 
      BackColor       =   &H80000005&
      Height          =   200
      Left            =   140
      TabIndex        =   2
      Top             =   2180
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   180
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
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "用于区别不同用途密码的简短代号, 如:淘宝账号可用 taobao 或 tb 等"
      Top             =   1463
      Width           =   1725
   End
   Begin VB.Timer tmrZip 
      Interval        =   1000
      Left            =   3000
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
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "可选择一个简单易记的密码, 用于生成其他高强度密码"
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
      ToolTipText     =   "用于区别不同用途密码的简短代号, 如:淘宝账号可用 taobao 或 tb 等"
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
      ToolTipText     =   "可选择一个简单易记的密码, 用于生成其他高强度密码"
      Top             =   877
      Width           =   3760
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "复制成功!"
      ForeColor       =   &H00CE9A02&
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblUserCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "附加扰码"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image iClose 
      Height          =   195
      Left            =   4080
      Picture         =   "FrmMain.frx":64E1
      Top             =   120
      Width           =   195
   End
   Begin VB.Image iSetting 
      Height          =   315
      Left            =   3600
      Picture         =   "FrmMain.frx":6567
      Top             =   2640
      Width           =   315
   End
   Begin VB.Image iInfo 
      Height          =   315
      Left            =   3240
      Picture         =   "FrmMain.frx":667E
      Top             =   2640
      Width           =   330
   End
   Begin VB.Image iHome 
      Height          =   315
      Left            =   2880
      Picture         =   "FrmMain.frx":6744
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label lblCode16 
      Alignment       =   2  'Center
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
      MouseIcon       =   "FrmMain.frx":680B
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "点击复制生成的十六位码"
      Top             =   1845
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Image imgInfo 
      Height          =   315
      Left            =   3210
      MouseIcon       =   "FrmMain.frx":695D
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":6AB3
      ToolTipText     =   "帮助"
      Top             =   2070
      Width           =   330
   End
   Begin VB.Image imgHome 
      Height          =   315
      Left            =   2842
      MouseIcon       =   "FrmMain.frx":6C64
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":6DBA
      ToolTipText     =   "官网"
      Top             =   2070
      Width           =   315
   End
   Begin VB.Image imgSetting 
      Height          =   315
      Left            =   3610
      MouseIcon       =   "FrmMain.frx":7036
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":718C
      ToolTipText     =   "设置"
      Top             =   2070
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   195
      Left            =   3720
      MouseIcon       =   "FrmMain.frx":7443
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":7599
      ToolTipText     =   "关闭"
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
' Copyright (c) 2011-2012 FlowerPassword.com All rights reserved.
'      Author : xLsDg @ Xiao Lu Software Development Group
'        Blog : http://hi.baidu.com/xlsdg
'          QQ : 4 4 7 4 0 5 7 4 0
'     Version : 1 . 1 . 0 . 0
'        Date : 2 0 1 2 / 0 4 / 1 2
' Description :
'     History :
'*****************************************************************
Option Explicit

Public Function calcKeyCode() As String

    Dim tmpKeyCode As String

    tmpKeyCode = comKey.Text

    If isPrefix Then
        tmpKeyCode = FrmSetting.comPrefixCode.Text + tmpKeyCode

    End If

    If isSuffix Then
        tmpKeyCode = tmpKeyCode + FrmSetting.comSuffixCode.Text

    End If

    calcKeyCode = tmpKeyCode

End Function

Public Sub FrmHide(ByVal isPostKey As Boolean, ByVal isKeyPress As Boolean)

    If isAlwaysOnTop Then UnSetWinOnTop Me.hwnd
    If isUseMouseHook Then Call UnSetMouseHook
    'Me.Visible = False
    Me.Hide

    If isPostKey Then
        If Len(lblCode16.Caption) > 0 Then
            PoseCodeToClipboard lblCode16.Caption

            If isKeyPress Then
                PostCode lblCode16.Caption, Password_Hwnd
            Else
                PostCode lblCode16.Caption, 0

            End If

            AddStrToCombox comKey, comKey.Text

        End If

    Else

        If Len(lblCode16.Caption) > 0 Then
            PoseCodeToClipboard lblCode16.Caption

        End If

    End If

    txtPassword.Text = ""

End Sub

Public Sub FrmShow()

    If isUseMouseHook Then
        If SetMouseHook = 0 Then
            MsgBox "花密设置全局鼠标按键挂钩失败!", vbCritical + vbOKOnly + vbSystemModal

        End If

    End If

    If isAlwaysOnTop Then
        If SetWinOnTop(Me.hwnd) = 0 Then
            MsgBox "花密设置主面板始终保持在其他窗口前端失败!", vbCritical + vbOKOnly + vbSystemModal

        End If

    End If

    'Me.Visible = True
    Me.Show

    If txtPassword.Enabled Then txtPassword.SetFocus

End Sub

Public Sub MidScreenShow()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Call FrmShow

End Sub

Public Sub ShowFrmByHotKey()

    If FrmSetting.Visible Then
        MsgBox "请先关闭系统设置窗口, 再尝试按全局热键显示花密!", vbCritical + vbOKOnly + vbSystemModal
    Else

        If Not Me.Visible Then

            Dim X As Long, Y As Long

            getLocation X, Y

            If isAutoUseDomain Then
                SetUrlAsKey Password_Hwnd

            End If

            Me.Move X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
            Call FrmShow

        End If

    End If

End Sub

Public Sub ShowResult()
    lblUserCode.ToolTipText = calcKeyCode
    If Len(txtPassword.Text) > 0 And Len(comKey.Text) > 0 Then
        If isShowCode Then lblCode16.Visible = True
        lblCode16.Caption = calcFlowerPassword(txtPassword.Text, calcKeyCode, calcPasswordLength(comKey.Text))
    Else
        lblCode16.Visible = False
        lblCode16.Caption = ""

    End If

End Sub

Private Sub ShowPasswordStrength()
    Dim lngLength As Long
    
    lngLength = Len(txtPassword.Text)
    If lngLength = 0 Then
        txtPassword.ToolTipText = "可选择一个简单易记的密码, 用于生成其他高强度密码"
        txtPassword.BackColor = &HFFFFFF
    Else

        Dim strHelp As String, strText As String

        Select Case check_password_level(txtPassword.Text, strHelp)

            Case PASSWORD_WEAK
                strText = "弱"
                txtPassword.BackColor = &HDEDEF2

            Case PASSWORD_NORMAL
                strText = "一般"
                txtPassword.BackColor = &HDAB23E

            Case PASSWORD_STRONG
                strText = "强"
                txtPassword.BackColor = &H1CBB3A '&HCE9A02

        End Select

        txtPassword.ToolTipText = strHelp
        'lblStrength.Caption = strText + " (" + CStr(lngLength) + ")"

    End If

End Sub

Private Sub chkAddUserCode_Click()

    If chkAddUserCode.value = Checked Then
        isAutoAddUserCode = True
        FrmSetting.chkAutoAddUserCode.value = Checked
    Else
        isAutoAddUserCode = False
        FrmSetting.chkAutoAddUserCode.value = Unchecked

    End If

    Call ShowResult

End Sub

Private Sub comKey_Change()

    If Len(comKey.Text) > 0 Then
        picKey.Visible = False
    Else
        picKey.Visible = True

    End If

    cbBox_Change comKey
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

Private Sub comKey_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        FrmHide False, False
    ElseIf KeyCode = vbKeyReturn Then
        FrmHide True, True
    ElseIf Len(comKey.Text) > 0 Then
        picKey.Visible = False

    End If

    cbBox_KeyDown comKey, KeyCode, Shift

End Sub

Private Sub comKey_LostFocus()

    If Len(comKey.Text) > 0 Then
        picKey.Visible = False
    Else
        picKey.Visible = True

    End If

    cbBox_LostFocus comKey

End Sub

Private Sub Form_Initialize()
    Me.ScaleMode = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
        FrmHide False, False

    End If

End Sub

Private Sub Form_Load()
    picPassword.Print "请输入记忆密码"
    picKey.Print "请输入区分代号"
    iClose.Left = imgClose.Left: iClose.Top = imgClose.Top
    iHome.Left = imgHome.Left: iHome.Top = imgHome.Top
    iInfo.Left = imgInfo.Left: iInfo.Top = imgInfo.Top
    iSetting.Left = imgSetting.Left: iSetting.Top = imgSetting.Top
    Call LoadSetting

    If isRunInIDEMode Then
        tmrZip.Enabled = False
    Else

        If App.PrevInstance Then
            MsgBox "花密已经运行, 请尝试结束后台进程再重新运行花密!", vbCritical + vbOKOnly + vbSystemModal
            End
        Else
            App.TaskVisible = False
            Call InitCommonControlsVB

            If isProtect Then
                If ProtectTextBox(txtPassword.hwnd) = 0 Then
                    MsgBox "花密设置记忆密码输入框保护措施失败!", vbCritical + vbOKOnly + vbSystemModal

                End If

            End If

            Call ZipMemory

        End If

    End If

    If SetHotKey(Me.hwnd) = 0 Then
        MsgBox "花密注册系统全局热键 (" + getHotKeyText + ") 失败!", vbCritical + vbOKOnly + vbSystemModal

    End If

    If AddToTray(FrmMain, App.Title) = 0 Then
        MsgBox "花密将图标添加至任务栏通知区域失败!", vbCritical + vbOKOnly + vbSystemModal
        End
    Else

        If isShowHelp Then
            SetTrayMsgbox "按全局热键【" + getHotKeyText + "】显示花密, 在主面板按 Enter 键或 Esc 键隐藏花密.", NIIF_GUID, "欢迎使用花密 Windows 桌面版", Me.Icon

        End If

    End If

    Call SetFormRgn

    If Not isAutoMini Then
        Call MidScreenShow

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetWinMove Me.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iClose.Visible = True: iHome.Visible = True: iInfo.Visible = True: iSetting.Visible = True: lblCopy.Visible = False

    If Not Me.Visible Then
        MouseOnTray Button, Shift, X, Y

    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If txtPassword.Enabled Then txtPassword.SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call UnSetFormRgn
    UnSetHotKey Me.hwnd

End Sub

Private Sub Form_Terminate()
    Call RemoveFromTray

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not isRunInIDEMode And isProtect Then
        UnProtectTextBox txtPassword.hwnd

    End If

End Sub

Private Sub iClose_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    iClose.Visible = False

End Sub

Private Sub iHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iHome.Visible = False

End Sub

Private Sub iInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iInfo.Visible = False

End Sub

Private Sub imgClose_Click()
    iClose.Visible = True
    FrmHide False, False

    If Not isCloseToHide Then
        Unload FrmSetting
        Unload Me

    End If

End Sub

Private Sub imgHome_Click()
    iHome.Visible = True
    OpenWebsite "http://flowerpassword.com"

End Sub

Private Sub imgInfo_Click()
    iInfo.Visible = True
    OpenWebsite "http://code.google.com/p/flower-password-windows"

End Sub

Private Sub imgSetting_Click()
    iSetting.Visible = True
    FrmHide False, False
    FrmSetting.Show

End Sub

Private Sub iSetting_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    iSetting.Visible = False

End Sub

Private Sub lblCode16_Click()

    If Len(lblCode16.Caption) > 0 Then
        Clipboard.Clear
        Clipboard.SetText lblCode16.Caption
        lblCopy.Visible = True

    End If

End Sub

Private Sub lblUserCode_Click()

    If chkAddUserCode.value = Checked Then
        chkAddUserCode.value = Unchecked
    Else
        chkAddUserCode.value = Checked

    End If

End Sub

Private Sub picKey_Click()

    'picKey.Visible = False
    If comKey.Enabled Then comKey.SetFocus

End Sub

Private Sub picPassword_Click()

    'picPassword.Visible = False
    If txtPassword.Enabled Then txtPassword.SetFocus

End Sub

Private Sub tmrZip_Timer()

    Static i As Integer

    i = i + 1

    If i > 60 Then
        i = 0
        Call ZipMemory

    End If

End Sub

Private Sub txtPassword_Change()

    If Len(txtPassword.Text) > 0 Then
        picPassword.Visible = False
    Else
        picPassword.Visible = True

    End If

    Call ShowPasswordStrength
    Call ShowResult

End Sub

Private Sub txtPassword_GotFocus()

    If Len(txtPassword.Text) > 0 Then
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)

        'Else
        'picPassword.Visible = False
    End If

End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        FrmHide False, False
    ElseIf KeyCode = vbKeyReturn Then

        If Len(comKey.Text) > 0 Then
            FrmHide True, True
        Else

            If comKey.Enabled Then comKey.SetFocus

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

