VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flower Password v1.0 build 20120311"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
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
   Picture         =   "FrmMain.frx":43B2
   ScaleHeight     =   6015
   ScaleWidth      =   9975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TmrShow 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5040
      Top             =   4800
   End
   Begin VB.TextBox txtCode32 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
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
      Left            =   1835
      Locked          =   -1  'True
      MaxLength       =   32
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   4285
      Width           =   3975
   End
   Begin VB.TextBox txtCode16 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
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
      Left            =   1835
      Locked          =   -1  'True
      MaxLength       =   16
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   3710
      Width           =   1960
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6490
      TabIndex        =   1
      Top             =   1060
      Width           =   3160
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
      Top             =   1060
      Width           =   3135
   End
   Begin VB.Image ImgLogo 
      Height          =   1920
      Left            =   7680
      Picture         =   "FrmMain.frx":5F1A
      Top             =   3720
      Width           =   1920
   End
   Begin VB.Image ImgCode 
      Height          =   360
      Left            =   1690
      MouseIcon       =   "FrmMain.frx":751E
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":7674
      ToolTipText     =   "Copy!!!"
      Top             =   3650
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label lblFlowerPassword 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2280
      MouseIcon       =   "FrmMain.frx":77C1
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5600
      Width           =   2010
   End
   Begin VB.Image ImgCopy 
      Height          =   360
      Left            =   5040
      Picture         =   "FrmMain.frx":7917
      Top             =   3645
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblCopy 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   3840
      MouseIcon       =   "FrmMain.frx":79F6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Copy!"
      Top             =   3640
      Width           =   900
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgCode.Visible = False
    
End Sub

Private Sub ImgCode_Click()

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

Private Sub lblFlowerPassword_Click()
    ShellExecute Me.hwnd, "Open", "http://flowerpassword.com/", 0, 0, 0

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
