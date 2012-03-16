VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flower Password v1.2 build 20120314"
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
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Public cMD5 As clsMD5

Private Sub ImgDonation_Click()
    ShellExecute Me.hwnd, "Open", "http://kisexu.com/go/huamidonation", 0, 0, 0

End Sub

Private Sub ImgJohnnyJian_Click()
    ShellExecute Me.hwnd, "Open", "http://johnnyjian.iteye.com/", 0, 0, 0

End Sub

Private Sub ImgKenshin_Click()
    ShellExecute Me.hwnd, "Open", "http://www.k-zone.cn/zblog/", 0, 0, 0

End Sub

Private Sub Imgkise_Click()
    ShellExecute Me.hwnd, "Open", "http://kisexu.com/", 0, 0, 0

End Sub

Private Sub Imgxlsdg_Click()
    ShellExecute Me.hwnd, "Open", "http://hi.baidu.com/xlsdg", 0, 0, 0

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
    ShellExecute Me.hwnd, "Open", "http://flowerpassword.com/", 0, 0, 0

End Sub

Private Sub PicContrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub txtCode32_GotFocus()
    txtCode32.SelStart = 0
    txtCode32.SelLength = Len(txtCode32.Text)

End Sub

Private Sub txtKey_Change()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        Call calFlowerPassword
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
        Call calFlowerPassword
    Else
        txtCode16.Text = ""
        txtCode32.Text = ""

    End If

End Sub

Private Sub calFlowerPassword()
    Set cMD5 = New clsMD5

    Dim md5one   As String

    Dim md5two   As String

    Dim md5three As String

    md5one = LCase$(cMD5.Hmac_MD5(txtPassword.Text, txtKey.Text))
    md5two = LCase$(cMD5.Hmac_MD5(md5one, "snow"))
    md5three = LCase$(cMD5.Hmac_MD5(LCase$(md5one), "kise"))

    Dim code32 As String

    code32 = ""

    Dim i As Integer

    For i = 1 To Len(md5two)

        If Not IsNumeric(Mid$(md5two, i, 1)) Then

            Dim str As String

            str = "sunlovesnow1990090127xykab"

            If InStr(1, str, Mid$(md5three, i, 1), vbBinaryCompare) > 0 Then
                code32 = code32 + UCase(Mid$(md5two, i, 1))
            Else
                code32 = code32 + Mid$(md5two, i, 1)

            End If

        Else
            code32 = code32 + Mid$(md5two, i, 1)

        End If

    Next

    Dim code1  As String

    Dim code16 As String

    code1 = Left$(code32, 1)

    If Not IsNumeric(code1) Then
        code16 = Left$(code32, 16)
    Else
        code16 = "K" + Mid$(code32, 2, 15)

    End If

    txtCode16.Text = code16
    txtCode32.Text = code32
    
    Set cMD5 = Nothing

End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
    
End Sub


