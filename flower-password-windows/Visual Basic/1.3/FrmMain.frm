VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flower Password v1.3 build 20120314"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9900
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicContrl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00F2F2F2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   0
      Picture         =   "FrmMain.frx":43B2
      ScaleHeight     =   4770
      ScaleWidth      =   9960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2340
      Width           =   9960
      Begin VB.Timer TmrShow 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   6840
         Top             =   3600
      End
      Begin VB.TextBox txtCode16 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FBFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1805
         Locked          =   -1  'True
         MaxLength       =   16
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   3700
         Width           =   1960
      End
      Begin VB.TextBox txtCode32 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FBFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1700
         Locked          =   -1  'True
         MaxLength       =   32
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   4300
         Width           =   3975
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
         Top             =   1058
         Width           =   3135
      End
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
         Height          =   360
         Left            =   6510
         TabIndex        =   1
         Top             =   1058
         Width           =   3120
      End
      Begin VB.Image ImgDonation 
         Height          =   585
         Left            =   8280
         MouseIcon       =   "FrmMain.frx":5699
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":57EB
         ToolTipText     =   "了解并资助花密的发展"
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Image Imgkise 
         Height          =   300
         Left            =   8880
         MouseIcon       =   "FrmMain.frx":5AC0
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":5C16
         ToolTipText     =   "徐小花"
         Top             =   4200
         Width           =   300
      End
      Begin VB.Image ImgKenshin 
         Height          =   300
         Left            =   8400
         MouseIcon       =   "FrmMain.frx":5F14
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":606A
         ToolTipText     =   "Kenshin"
         Top             =   4200
         Width           =   300
      End
      Begin VB.Image ImgJohnnyJian 
         Height          =   300
         Left            =   7920
         MouseIcon       =   "FrmMain.frx":63AB
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":6501
         ToolTipText     =   "JohnnyJian"
         Top             =   4200
         Width           =   300
      End
      Begin VB.Image Imgxlsdg 
         Height          =   300
         Left            =   9360
         MouseIcon       =   "FrmMain.frx":66DC
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":6832
         ToolTipText     =   "xLsDg"
         Top             =   4200
         Width           =   300
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2011-2012 FlowerPassword.com All rights reserved."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5280
         MouseIcon       =   "FrmMain.frx":6CB7
         MousePointer    =   99  'Custom
         TabIndex        =   6
         ToolTipText     =   "http://flowerpassword.com/"
         Top             =   120
         Width           =   4500
      End
      Begin VB.Label lblCopy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   345
         Left            =   3820
         MouseIcon       =   "FrmMain.frx":6E0D
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   3640
         Width           =   900
      End
      Begin VB.Image ImgCode 
         Height          =   360
         Left            =   1680
         MouseIcon       =   "FrmMain.frx":6F63
         MousePointer    =   99  'Custom
         Picture         =   "FrmMain.frx":70B9
         Top             =   3640
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Image ImgCopy 
         Height          =   360
         Left            =   5040
         Picture         =   "FrmMain.frx":7206
         Top             =   3640
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Image ImgLogo 
      Height          =   2340
      Left            =   0
      Picture         =   "FrmMain.frx":72E5
      Top             =   0
      Width           =   9900
   End
   Begin VB.Menu munFiles 
      Caption         =   "文件(&F)"
      Visible         =   0   'False
      Begin VB.Menu munShow 
         Caption         =   "显示(&S)"
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

Private Const WM_LBUTTONUP = &H202

Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_RBUTTONUP = &H205

Private Const WM_RBUTTONDBLCLK = &H206

Dim blnHide As Boolean

Private Sub Form_Load()
    ProtectTextBox txtPassword.hwnd
    SetHotKey Me.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnProtectTextBox txtPassword.hwnd
    UnSetHotKey Me.hwnd

End Sub

Private Sub ImgDonation_Click()
    OpenWebsite "http://kisexu.com/go/huamidonation"

End Sub

Private Sub ImgJohnnyJian_Click()
    OpenWebsite "http://johnnyjian.iteye.com"

End Sub

Private Sub ImgKenshin_Click()
    OpenWebsite "http://www.k-zone.cn/zblog"

End Sub

Private Sub Imgkise_Click()
    OpenWebsite "http://kisexu.com"

End Sub

Private Sub ImgLogo_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)

    If (blnHide) Then

        Dim lMsg As Single

        lMsg = x / Screen.TwipsPerPixelX

        Select Case lMsg

            Case WM_LBUTTONUP                   '左键抬起
                Me.WindowState = vbNormal
                Me.Show
                blnHide = False
                Call RemoveFromTray           '删除托盘图标

            Case WM_LBUTTONDBLCLK               '左键双击

            Case WM_RBUTTONUP                   '右键抬起
                PopupMenu munFiles

            Case WM_RBUTTONDBLCLK               '右键双击

                'Case WM_MBUTTONUP                   '中键抬起
                'Case WM_MBUTTONDBLCLK               '中键双击
            Case 1028                           '点击托盘气泡关闭按钮

            Case 1029                           '点击托盘气泡窗体本身

                'If blnLostNet Then
                '    blnLostNet = False
                '    Call cmdLogin_Click
                'End If
        End Select

    End If

End Sub

Private Sub Imgxlsdg_Click()
    OpenWebsite "http://hi.baidu.com/xlsdg"

End Sub

Private Sub lblCopy_Click()

    If Len(txtCode16.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText txtCode16
        TmrShow.Enabled = True
        ImgCopy.Visible = True

    End If

End Sub

Private Sub lblCopy_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)
    ImgCode.Visible = True

End Sub

Private Sub lblCopyright_Click()
    OpenWebsite "http://flowerpassword.com"

End Sub

Private Sub munShow_Click()
    Me.WindowState = vbNormal
    Me.Show
    blnHide = False
    Call RemoveFromTray           '删除托盘图标

End Sub

Private Sub PicContrl_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    ImgCode.Visible = False

End Sub

Private Sub TmrShow_Timer()

    Static count As Long

    count = count + 1

    If count > 5 Then
        count = 0
        TmrShow.Enabled = False
        ImgCopy.Visible = False

    End If

End Sub

Private Sub txtCode16_GotFocus()
    txtCode16.SelStart = 0
    txtCode16.SelLength = Len(txtCode16.Text)

End Sub

Private Sub txtCode16_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    ImgCode.Visible = False

End Sub

Private Sub txtCode32_GotFocus()
    txtCode32.SelStart = 0
    txtCode32.SelLength = Len(txtCode32.Text)

End Sub

Private Sub txtKey_Change()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        txtCode16.Text = getFlowerPassword(txtPassword.Text, txtKey.Text, 16)
        txtCode32.Text = getFlowerPassword(txtPassword.Text, txtKey.Text, 32)
    Else
        txtCode16.Text = ""
        txtCode32.Text = ""

    End If

End Sub

Private Sub txtKey_GotFocus()
    txtKey.SelStart = 0
    txtKey.SelLength = Len(txtKey.Text)

End Sub

Private Sub txtPassword_Change()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        txtCode16.Text = getFlowerPassword(txtPassword.Text, txtKey.Text, 16)
        txtCode32.Text = getFlowerPassword(txtPassword.Text, txtKey.Text, 32)
    Else
        txtCode16.Text = ""
        txtCode32.Text = ""

    End If

End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        Me.Hide
        AddToTray FrmMain, App.Title
        blnHide = True

    End If

End Sub

Private Sub munExit_Click()
    Call RemoveFromTray
    Unload Me

End Sub
