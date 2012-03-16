VERSION 5.00
Begin VB.Form FrmInput 
   BorderStyle     =   0  'None
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox PicUI 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3720
      Left            =   0
      Picture         =   "FrmInput.frx":0000
      ScaleHeight     =   3720
      ScaleWidth      =   4110
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   4110
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
         Top             =   690
         Width           =   2415
      End
      Begin VB.Label lblClose 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   3800
         MouseIcon       =   "FrmInput.frx":16DB
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   315
      End
      Begin VB.Label lblWebsite 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   1530
         MouseIcon       =   "FrmInput.frx":1831
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Open"
         Top             =   3270
         Width           =   2385
      End
   End
End
Attribute VB_Name = "FrmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Declare Function SendMessage _
                Lib "user32.dll" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const WM_SYSCOMMAND = &H112

Private Const SC_MOVE = &HF010&

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Private Sub Form_Load()
    SetWinOnTop Me.hwnd
    ProtectTextBox txtPassword.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnSetWinOnTop Me.hwnd
    UnProtectTextBox txtPassword.hwnd

End Sub

Private Sub lblClose_Click()
    Me.Visible = False

End Sub

Private Sub lblWebsite_Click()
    OpenWebsite "http://flowerpassword.com"

End Sub

Private Sub PicUI_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Or KeyCode = 13 Then
        Me.Visible = False

    End If

End Sub

Private Sub PicUI_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0

End Sub

Private Sub txtKey_Change()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText getFlowerPassword(txtPassword.Text, txtKey.Text, 16)

    End If

End Sub

Private Sub txtKey_GotFocus()
    txtKey.SelStart = 0
    txtKey.SelLength = Len(txtKey.Text)

End Sub

Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Or KeyCode = 13 Then
        Me.Visible = False

    End If

End Sub

Private Sub txtPassword_Change()

    If Len(txtPassword.Text) > 0 And Len(txtKey.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText getFlowerPassword(txtPassword.Text, txtKey.Text, 16)

    End If

End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Me.Visible = False
    ElseIf KeyCode = 13 Then
        txtKey.SetFocus

    End If

End Sub
