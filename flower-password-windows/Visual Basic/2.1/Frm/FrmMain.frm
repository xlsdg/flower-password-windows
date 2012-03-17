VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flower Password"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
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
   ScaleHeight     =   1965
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   1
      ToolTipText     =   "用于区别不同用途密码的简短代号，如淘宝账号可用“taobao”或“tb”等。"
      Top             =   1290
      Width           =   2415
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "可选择一个简单易记的密码，用于生成其他高强度密码。(建议：选择一个与个人信息无关的密码，防止社会工程学破解。)"
      Top             =   690
      Width           =   2415
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   3800
      MouseIcon       =   "FrmMain.frx":4BDD
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   315
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   350
      Left            =   910
      MouseIcon       =   "FrmMain.frx":4D33
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Go to FlowerPassword.com"
      Top             =   160
      Width           =   2745
   End
   Begin VB.Menu munFlowerPassword 
      Caption         =   "花密(&F)"
      Visible         =   0   'False
      Begin VB.Menu munSetting 
         Caption         =   "设置(&S)"
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

Private Sub Form_Initialize()

    If App.PrevInstance Then
        MsgBox "花密已经运行，请关闭后台程序再重试一次！", vbCritical + vbOKOnly
        End
    Else
        App.TaskVisible = False
        InitCommonControlsVB

        If ProtectTextBox(txtPassword.hwnd) = 0 Then
            MsgBox "花密记忆密码输入框保护失败！", vbCritical + vbOKOnly

        End If

    End If

End Sub

Private Sub Form_Load()
    SetWinStyle Me.hwnd: Me.Width = 4110: Me.Height = 1965
    SetWinOnTop Me.hwnd

    If SetHotKey(Me.hwnd) = 0 Then
        MsgBox "花密快捷键（WIN + S）与系统其它程序冲突！", vbCritical + vbOKOnly
        End

    End If

    If SetMouseHook = 0 Then
        MsgBox "花密加载鼠标挂钩模块失败！", vbCritical + vbOKOnly
        End

    End If

    If AddToTray(FrmMain, App.Title) = 0 Then
        MsgBox "花密加载托盘模块失败！", vbCritical + vbOKOnly
        End
    Else
        SetTrayMsgbox "按快捷键 WIN + S 呼唤花密，按Enter键或Esc键关闭输入窗口。", NIIF_GUID, "欢迎您使用花密Windows版", Me.Icon

    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not Me.Visible Then
        MouseOnTray Button, Shift, X, Y

    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnSetMouseHook
    UnSetHotKey Me.hwnd
    UnSetWinOnTop Me.hwnd

End Sub

Private Sub munExit_Click()
    Unload Me

End Sub

Private Sub Form_Terminate()
    RemoveFromTray

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnProtectTextBox txtPassword.hwnd

End Sub

Private Sub lblClose_Click()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText getFlowerPassword(txtPassword.Text, txtKey.Text, 16)

    End If

    txtPassword.Text = ""
    Me.Visible = False

End Sub

Private Sub lblWebsite_Click()
    OpenWebsite "http://flowerpassword.com"

End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)

    If keycode = 27 Or keycode = 13 Then
        If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
            Clipboard.Clear
            Clipboard.SetText getFlowerPassword(txtPassword.Text, txtKey.Text, 16)

        End If

        txtPassword.Text = ""
        Me.Visible = False

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetWinMove Me.hwnd

End Sub

Private Sub munSetting_Click()
    MsgBox "==(^。^)== 期待...", vbInformation + vbOKOnly

End Sub

Private Sub txtKey_GotFocus()
    txtKey.SelStart = 0
    txtKey.SelLength = Len(txtKey.Text)

End Sub

Private Sub txtKey_KeyDown(keycode As Integer, Shift As Integer)

    If keycode = 27 Then
        txtPassword.Text = ""
        Me.Visible = False
    ElseIf keycode = 13 Then

        If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
            PostCode Password_Hwnd, getFlowerPassword(txtPassword.Text, txtKey.Text, 16)

        End If

        txtPassword.Text = ""
        Me.Visible = False

    End If

End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

Private Sub txtPassword_KeyDown(keycode As Integer, Shift As Integer)

    If keycode = 27 Then
        txtPassword.Text = ""
        Me.Visible = False
    ElseIf keycode = 13 Then

        If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
            Clipboard.Clear
            Clipboard.SetText getFlowerPassword(txtPassword.Text, txtKey.Text, 16)
            txtPassword.Text = ""
            Me.Visible = False
        Else
            txtKey.SetFocus

        End If

    End If

End Sub
