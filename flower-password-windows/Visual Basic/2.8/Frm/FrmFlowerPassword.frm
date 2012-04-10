VERSION 5.00
Begin VB.Form FrmFlowerPassword 
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flower Password v2.8 build 20120407"
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
   Icon            =   "FrmFlowerPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox PicContrl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00F2F2F2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   0
      Picture         =   "FrmFlowerPassword.frx":43B2
      ScaleHeight     =   4770
      ScaleWidth      =   9960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2340
      Width           =   9960
      Begin VB.CheckBox chkAutoHide 
         BackColor       =   &H00FFFFFF&
         Caption         =   "下次启动时自动最小化至系统托盘（提示：双击托盘图标可重新打开花密主界面）"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4200
         Width           =   6855
      End
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
         MouseIcon       =   "FrmFlowerPassword.frx":5699
         MousePointer    =   99  'Custom
         Picture         =   "FrmFlowerPassword.frx":57EB
         ToolTipText     =   "了解并资助花密的发展"
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Image Imgkise 
         Height          =   300
         Left            =   8880
         MouseIcon       =   "FrmFlowerPassword.frx":5AC0
         MousePointer    =   99  'Custom
         Picture         =   "FrmFlowerPassword.frx":5C16
         ToolTipText     =   "KiseXu"
         Top             =   4200
         Width           =   300
      End
      Begin VB.Image ImgKenshin 
         Height          =   300
         Left            =   8400
         MouseIcon       =   "FrmFlowerPassword.frx":5F14
         MousePointer    =   99  'Custom
         Picture         =   "FrmFlowerPassword.frx":606A
         ToolTipText     =   "Kenshin"
         Top             =   4200
         Width           =   300
      End
      Begin VB.Image ImgJohnnyJian 
         Height          =   300
         Left            =   7920
         MouseIcon       =   "FrmFlowerPassword.frx":63AB
         MousePointer    =   99  'Custom
         Picture         =   "FrmFlowerPassword.frx":6501
         ToolTipText     =   "JohnnyJian"
         Top             =   4200
         Width           =   300
      End
      Begin VB.Image Imgxlsdg 
         Height          =   300
         Left            =   9360
         MouseIcon       =   "FrmFlowerPassword.frx":66DC
         MousePointer    =   99  'Custom
         Picture         =   "FrmFlowerPassword.frx":6832
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
         MouseIcon       =   "FrmFlowerPassword.frx":6CB7
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "http://flowerpassword.com/"
         Top             =   120
         Width           =   4500
      End
      Begin VB.Label lblCopy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   345
         Left            =   3820
         MouseIcon       =   "FrmFlowerPassword.frx":6E0D
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3640
         Width           =   900
      End
      Begin VB.Image ImgCode 
         Height          =   360
         Left            =   1680
         MouseIcon       =   "FrmFlowerPassword.frx":6F63
         MousePointer    =   99  'Custom
         Picture         =   "FrmFlowerPassword.frx":70B9
         Top             =   3640
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Image ImgCopy 
         Height          =   360
         Left            =   5040
         Picture         =   "FrmFlowerPassword.frx":7206
         Top             =   3640
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Image ImgLogo 
      Height          =   2340
      Left            =   0
      Picture         =   "FrmFlowerPassword.frx":72E5
      Top             =   0
      Width           =   9900
   End
End
Attribute VB_Name = "FrmFlowerPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
' Copyright (c) 2011-2012 FlowerPassword.com All rights reserved.
'      Author : xLsDg @ Xiao Lu Software Development Group
'        Blog : http://hi.baidu.com/xlsdg
'          QQ : 4 4 7 4 0 5 7 4 0
'     Version : 1 . 0 . 0 . 0
'        Date : 2 0 1 2 / 0 4 / 0 7
' Description :
'     History :
'*****************************************************************
Option Explicit

Private Sub chkAutoHide_Click()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    If chkAutoHide.value = Checked Then
        isAutoMini = True
        WriteIni "Setting", "AutoMini", "1", strSettingPath
    Else
        isAutoMini = False
        WriteIni "Setting", "AutoMini", "0", strSettingPath

    End If

End Sub

Private Sub Form_Initialize()

    If Not isRunInIDEMode Then
        If ProtectTextBox(txtPassword.hwnd) = 0 Then
            MsgBox "花密记忆密码输入框保护失败！", vbCritical + vbOKOnly + vbSystemModal

        End If

    End If

    Call InitCommonControlsVB

End Sub

Private Sub Form_Load()

    If isAutoMini Then
        chkAutoHide.value = Checked
    Else
        chkAutoHide.value = Unchecked

    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Not isExit Then
        If MsgBox("选择[ 是 - Yes ]直接退出花密，选择[ 否 - No ]最小化花密至系统托盘", vbQuestion + vbYesNo + vbSystemModal) = vbNo Then
            Cancel = True
            isExit = False
            Me.Hide
        Else
            isExit = True
            Unload Me
            Unload FrmMain
    
        End If
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not isRunInIDEMode Then UnProtectTextBox txtPassword.hwnd

End Sub

Private Sub ImgDonation_Click()
    OpenWebsite "http://kisexu.com/go/huamidonation"

End Sub

Private Sub ImgJohnnyJian_Click()
    OpenWebsite "http://johnnyjian.iteye.com/"

End Sub

Private Sub ImgKenshin_Click()
    OpenWebsite "http://www.k-zone.cn/zblog/"

End Sub

Private Sub Imgkise_Click()
    OpenWebsite "http://kisexu.com/"

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
                              X As Single, _
                              Y As Single)
    ImgCode.Visible = True

End Sub

Private Sub lblCopyright_Click()
    OpenWebsite "http://flowerpassword.com/"

End Sub

Private Sub PicContrl_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
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
                                X As Single, _
                                Y As Single)
    ImgCode.Visible = False

End Sub

Private Sub txtKey_Change()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        txtCode16.Text = getFlowerPassword(txtPassword.Text, txtKey.Text, 16)
    Else
        txtCode16.Text = ""

    End If

End Sub

Private Sub txtKey_GotFocus()
    txtKey.SelStart = 0
    txtKey.SelLength = Len(txtKey.Text)

End Sub

Private Sub txtPassword_Change()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        txtCode16.Text = getFlowerPassword(txtPassword.Text, txtKey.Text, 16)
    Else
        txtCode16.Text = ""

    End If

End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

