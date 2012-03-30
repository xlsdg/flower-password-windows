VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flower Password"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":43B2
   ScaleHeight     =   3615
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picSetting 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3150
      Left            =   4320
      Picture         =   "FrmMain.frx":6829
      ScaleHeight     =   3150
      ScaleWidth      =   4050
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Timer TmrZip 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox PicUser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      ScaleWidth      =   3300
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "附加扰码可以为空。"
      Top             =   2080
      Width           =   3300
   End
   Begin VB.PictureBox PicKey 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      ScaleWidth      =   3045
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "用于区别不同用途密码的简短代号，如淘宝账号可用“taobao”或“tb”等。"
      Top             =   1520
      Width           =   3045
   End
   Begin VB.PictureBox PicPassword 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      ScaleWidth      =   3285
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "可选择一个简单易记的密码，用于生成其他高强度密码。"
      Top             =   970
      Width           =   3285
   End
   Begin VB.TextBox txtUser 
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
      TabIndex        =   2
      ToolTipText     =   "附加扰码。"
      Top             =   2030
      Width           =   3760
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
      ToolTipText     =   "用于区别不同用途密码的简短代号，如淘宝账号可用“taobao”或“tb”等。"
      Top             =   1470
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
      ToolTipText     =   "可选择一个简单易记的密码，用于生成其他高强度密码。"
      Top             =   920
      Width           =   3760
   End
   Begin VB.Image ImgCopy 
      Height          =   360
      Left            =   1900
      Picture         =   "FrmMain.frx":70CA
      Top             =   2655
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image iClose 
      Height          =   195
      Left            =   4080
      Picture         =   "FrmMain.frx":7192
      Top             =   120
      Width           =   195
   End
   Begin VB.Image iSetting 
      Height          =   315
      Left            =   3600
      Picture         =   "FrmMain.frx":7218
      Top             =   3240
      Width           =   315
   End
   Begin VB.Image iInfo 
      Height          =   315
      Left            =   3240
      Picture         =   "FrmMain.frx":7342
      Top             =   3240
      Width           =   330
   End
   Begin VB.Image iHome 
      Height          =   315
      Left            =   2880
      Picture         =   "FrmMain.frx":7447
      Top             =   3240
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
      Left            =   150
      MouseIcon       =   "FrmMain.frx":7546
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "点击复制十六位码（推荐）！"
      Top             =   2715
      Width           =   60
   End
   Begin VB.Image imgInfo 
      Height          =   315
      Left            =   3200
      MouseIcon       =   "FrmMain.frx":7698
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":77EE
      ToolTipText     =   "帮助"
      Top             =   2685
      Width           =   330
   End
   Begin VB.Image imgHome 
      Height          =   315
      Left            =   2840
      MouseIcon       =   "FrmMain.frx":799F
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":7AF5
      ToolTipText     =   "主页"
      Top             =   2685
      Width           =   315
   End
   Begin VB.Image imgSetting 
      Height          =   315
      Left            =   3600
      MouseIcon       =   "FrmMain.frx":7D71
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":7EC7
      ToolTipText     =   "设置"
      Top             =   2680
      Width           =   315
   End
   Begin VB.Image ImgClose 
      Height          =   195
      Left            =   3720
      MouseIcon       =   "FrmMain.frx":817F
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":82D5
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

Public Sub AddStrToComKey()

    If Not CompareComKey(comKey.Text) Then
        comKey.AddItem comKey.Text

    End If

End Sub

Public Sub FrmHide()
    Call UnSendCodeToEditBox
    UnSetWinOnTop Me.hWnd
    Call UnSetMouseHook

End Sub

Public Sub FrmShow()

    If SetMouseHook = 0 Then
        MsgBox "花密加载鼠标挂钩模块失败！", vbCritical + vbOKOnly
        Unload Me

    End If

    Me.Visible = True
    
    If SetWinOnTop(Me.hWnd) = 0 Then
        MsgBox "花密窗体置顶失败！", vbCritical + vbOKOnly
        Unload Me
        
    End If
    
    txtPassword.SetFocus

End Sub

Public Sub SendCodeToEditBox(ByVal Key As Boolean)
    Me.Visible = False

    If Len(lblCode16.Caption) > 0 Then
        Clipboard.Clear
        Clipboard.SetText lblCode16.Caption

        If Key Then
            PostCode lblCode16.Caption, Password_Hwnd
        Else
            PostCode lblCode16.Caption, 0

        End If

        AddStrToComKey

    End If

    txtPassword.Text = ""

End Sub

Public Sub UnSendCodeToEditBox()
    Me.Visible = False

    If Len(lblCode16.Caption) > 0 Then
        Clipboard.Clear
        Clipboard.SetText lblCode16.Caption

    End If

    txtPassword.Text = ""

End Sub

Private Function CompareComKey(ByVal strKey As String) As Boolean

    Dim i As Long

    CompareComKey = False

    For i = 0 To comKey.ListCount - 1

        If comKey.List(i) = strKey Then
            CompareComKey = True
            Exit For

        End If

    Next

End Function

Private Sub comKey_Change()

    If Len(comKey.Text) > 0 Then
        PicKey.Visible = False
    Else
        PicKey.Visible = True

    End If

    If Len(comKey.Text) > 0 And Len(txtPassword.Text) > 0 Then
        lblCode16.Caption = getFlowerPassword(txtPassword.Text, comKey.Text + txtUser.Text, 16)
    Else
        lblCode16.Caption = ""

    End If

End Sub

Private Sub comKey_Click()

    If Len(comKey.Text) > 0 And Len(txtPassword.Text) > 0 Then
        lblCode16.Caption = getFlowerPassword(txtPassword.Text, comKey.Text + txtUser.Text, 16)
    Else
        lblCode16.Caption = ""

    End If

End Sub

Private Sub comKey_GotFocus()
    PicKey.Visible = False
    comKey.SelStart = 0
    comKey.SelLength = Len(comKey.Text)

End Sub

Private Sub comKey_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call FrmHide
    ElseIf KeyCode = 13 Then

        If Len(lblCode16.Caption) > 0 Then
            SendCodeToEditBox True
        Else
            txtUser.SetFocus

        End If

    ElseIf Len(comKey.Text) > 0 Then
        PicKey.Visible = False

    End If

End Sub

Private Sub comKey_LostFocus()

    If Len(comKey.Text) > 0 Then
        PicKey.Visible = False
    Else
        PicKey.Visible = True

    End If

End Sub

Private Sub Form_Initialize()

    If App.PrevInstance Then
        MsgBox "花密已经运行，请关闭后台程序再重试一次！", vbCritical + vbOKOnly
        End
    Else
        App.TaskVisible = False
        InitCommonControlsVB

        If ProtectTextBox(txtPassword.hWnd) = 0 Then
            MsgBox "花密记忆密码输入框保护失败！", vbCritical + vbOKOnly

        End If

    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Or KeyCode = 13 Then
        Call FrmHide

    End If

End Sub

Private Sub Form_Load()
    PicPassword.Print "请输入记忆密码": PicKey.Print "请输入区分代号": PicUser.Print "请输入附加扰码"
    SetWinStyle Me.hWnd: Me.Width = 4065: Me.Height = 3165
    iClose.Left = ImgClose.Left: iClose.Top = ImgClose.Top: iHome.Left = imgHome.Left: iHome.Top = imgHome.Top: iInfo.Left = imgInfo.Left: iInfo.Top = imgInfo.Top: iSetting.Left = imgSetting.Left: iSetting.Top = imgSetting.Top

    If SetHotKey(Me.hWnd) = 0 Then
        MsgBox "花密快捷键（WIN + S）与系统其它程序冲突！", vbCritical + vbOKOnly
        End

    End If

    If AddToTray(FrmMain, App.Title) = 0 Then
        MsgBox "花密加载托盘模块失败！", vbCritical + vbOKOnly
        End
    Else
        SetTrayMsgbox "按快捷键 WIN + S 呼唤花密，按Enter键或Esc键关闭输入窗口。", NIIF_GUID, "欢迎您使用花密Windows版", Me.Icon

    End If
    
    If SetFormRgn = 0 Then
        MsgBox "花密窗体特效设置失败！", vbCritical + vbOKOnly
        End
        
    End If
    
    Call ZipMemory

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetWinMove Me.hWnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iClose.Visible = True: iHome.Visible = True: iInfo.Visible = True: iSetting.Visible = True: ImgCopy.Visible = False

    If Not Me.Visible Then
        MouseOnTray Button, Shift, X, Y

    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnSetFormRgn = 0 Then
        MsgBox "花密窗体特效撤销失败！", vbCritical + vbOKOnly
        'End
        
    End If
    
    If UnSetHotKey(Me.hWnd) = 0 Then
        MsgBox "花密快捷键（WIN + S）撤销失败！", vbCritical + vbOKOnly
        'End
        
    End If

End Sub

Private Sub Form_Terminate()
    If RemoveFromTray = 0 Then
        MsgBox "花密卸载托盘模块失败！", vbCritical + vbOKOnly
        'End
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnProtectTextBox txtPassword.hWnd

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
    Call FrmHide
    iClose.Visible = True

End Sub

Private Sub imgHome_Click()
    OpenWebsite "http://flowerpassword.com"
    iHome.Visible = True

End Sub

Private Sub imgInfo_Click()
    OpenWebsite "http://flowerpassword.com/guide"
    iInfo.Visible = True

End Sub

Private Sub imgSetting_Click()
    iSetting.Visible = True
    picSetting.Left = 0: picSetting.Top = 0
    picSetting.Visible = True

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
        ImgCopy.Visible = True

    End If

End Sub

Private Sub munExit_Click()
    Unload Me

End Sub

Private Sub munShow_Click()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Call FrmShow

End Sub

Private Sub PicKey_Click()
    PicKey.Visible = False
    comKey.SetFocus

End Sub

Private Sub PicPassword_Click()
    PicPassword.Visible = False
    txtPassword.SetFocus

End Sub

Private Sub PicUser_Click()
    PicUser.Visible = False
    txtUser.SetFocus

End Sub

Private Sub TmrZip_Timer()

    Static i As Integer

    i = i + 1

    If i > 60 Then
        i = 0
        Call ZipMemory

    End If

End Sub

Private Sub txtPassword_Change()

    If Len(txtPassword.Text) > 0 Then
        PicPassword.Visible = False
    Else
        PicPassword.Visible = True

    End If

    If Len(txtPassword.Text) > 0 And Len(comKey.Text) > 0 Then
        lblCode16.Caption = getFlowerPassword(txtPassword.Text, comKey.Text + txtUser.Text, 16)
    Else
        lblCode16.Caption = ""

    End If

End Sub

Private Sub txtPassword_GotFocus()
    PicPassword.Visible = False
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call FrmHide
    ElseIf KeyCode = 13 Then

        If Len(lblCode16.Caption) > 0 And Len(comKey.Text) > 0 Then
            SendCodeToEditBox True
        Else
            comKey.SetFocus

        End If

    ElseIf Len(txtPassword.Text) > 0 Then
        PicPassword.Visible = False

    End If

End Sub

Private Sub txtPassword_LostFocus()

    If Len(txtPassword.Text) > 0 Then
        PicPassword.Visible = False
    Else
        PicPassword.Visible = True

    End If

End Sub

Private Sub txtUser_Change()

    If Len(txtUser.Text) > 0 Then
        PicUser.Visible = False
    Else
        PicUser.Visible = True

    End If

    If Len(txtPassword.Text) > 0 And Len(comKey.Text) > 0 Then
        lblCode16.Caption = getFlowerPassword(txtPassword.Text, comKey.Text + txtUser.Text, 16)
    Else
        lblCode16.Caption = ""

    End If

End Sub

Private Sub txtUser_GotFocus()
    PicUser.Visible = False
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser.Text)

End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call FrmHide
    ElseIf KeyCode = 13 Then

        If Len(lblCode16.Caption) > 0 Then
            SendCodeToEditBox True
        Else
            txtPassword.SetFocus

        End If

    ElseIf Len(txtUser.Text) > 0 Then
        PicUser.Visible = False

    End If

End Sub

Private Sub txtUser_LostFocus()

    If Len(txtUser.Text) > 0 Then
        PicUser.Visible = False
    Else
        PicUser.Visible = True

    End If

End Sub
