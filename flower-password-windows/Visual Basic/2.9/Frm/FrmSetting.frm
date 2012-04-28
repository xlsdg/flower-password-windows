VERSION 5.00
Begin VB.Form FrmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "系统设置"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDefault 
      Caption         =   "恢复默认设置"
      Height          =   360
      Left            =   240
      TabIndex        =   17
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame FraFunction 
      Caption         =   "功能和选项"
      Height          =   3615
      Left            =   240
      TabIndex        =   21
      Top             =   2880
      Width           =   4815
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Index           =   1
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4575
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   4575
         Begin VB.CheckBox chkUseMouseHook 
            Caption         =   "使用全局鼠标按键挂钩方式辅助定位密码输入框"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            ToolTipText     =   "针对使用模拟按键输入方式时, 当前输入状态不在密码输入框的情况"
            Top             =   3000
            Width           =   4215
         End
         Begin VB.PictureBox picUserCode 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   700
            MousePointer    =   3  'I-Beam
            ScaleHeight     =   300
            ScaleWidth      =   1725
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "自动在区分代号中添加附加扰码"
            Top             =   2563
            Width           =   1725
         End
         Begin VB.CheckBox chkAutoUseDomain 
            Caption         =   "自动将浏览器地址栏网站域名填入区分代号"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            ToolTipText     =   "密码输入框在浏览器窗口内部时, 自动获取地址栏网站域名作为区分代号"
            Top             =   360
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.CheckBox chkAutoAddUserCode 
            Caption         =   "自动保存并默认填入上次使用的附加扰码 :"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "自动将附加扰码添加到区分代号相应位置"
            Top             =   1440
            Width           =   3855
         End
         Begin VB.CheckBox chkAutoCopy 
            Caption         =   "自动将生成的密码复制到剪贴板"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   0
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.OptionButton optPrefix 
            Caption         =   "前缀 ( 附加扰码 + 区分代号 )"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   11
            Top             =   1800
            Width           =   3255
         End
         Begin VB.OptionButton optSuffix 
            Caption         =   "后缀 ( 区分代号 + 附加扰码 )"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   12
            Top             =   2160
            Width           =   3255
         End
         Begin VB.CheckBox chkDomainSuffix 
            Caption         =   "包含网站域名后缀 (如: .com , .net , .org , ...)"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   720
            Width           =   3915
         End
         Begin VB.ComboBox comUserCode 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   600
            TabIndex        =   13
            ToolTipText     =   "自动在区分代号中添加附加扰码"
            Top             =   2520
            Width           =   3720
         End
         Begin VB.CheckBox chkAutoCheckClipboard 
            Caption         =   "自动识别剪贴板复制的网址 (推荐)"
            Height          =   255
            Left            =   600
            TabIndex        =   9
            ToolTipText     =   "对于不能自动获取地址栏网站域名的浏览器, 您可以先手动复制URL至剪贴板由花密自动识别域名作为区分代号"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3255
         End
      End
   End
   Begin VB.Frame FraStartAndClose 
      Caption         =   "启动和关闭"
      Height          =   2655
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   4815
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Index           =   0
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   4575
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   4575
         Begin VB.CheckBox chkShowCode 
            Caption         =   "在花密主面板显示生成的十六位密码"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.CheckBox chkTransparent 
            Caption         =   "开启花密主面板半透明度效果"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   720
            Width           =   2895
         End
         Begin VB.CheckBox chkAlwaysOnTop 
            Caption         =   "花密主面板始终保持在其他窗口前端"
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.OptionButton OptCloseToExit 
            Caption         =   "退出程序"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   2040
            Width           =   1335
         End
         Begin VB.OptionButton OptClodeToHide 
            Caption         =   "隐藏到任务栏通知区域 , 不退出程序"
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   1680
            Value           =   -1  'True
            Width           =   3495
         End
         Begin VB.CheckBox chkAutoMini 
            Caption         =   "启动花密时为我自动最小化至任务栏通知区域"
            Height          =   255
            Left            =   240
            TabIndex        =   0
            Top             =   0
            Width           =   4215
         End
         Begin VB.Label lblAskClose 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "关闭花密主面板时 :"
            Height          =   195
            Left            =   480
            TabIndex        =   20
            Top             =   1410
            Width           =   1545
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   360
      Left            =   3960
      TabIndex        =   16
      Top             =   6840
      Width           =   1110
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   360
      Left            =   2640
      TabIndex        =   15
      Top             =   6840
      Width           =   1110
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5160
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Menu munFlowerPassword 
      Caption         =   "花密"
      Visible         =   0   'False
      Begin VB.Menu munShow 
         Caption         =   "显示"
      End
      Begin VB.Menu munLine 
         Caption         =   "-"
      End
      Begin VB.Menu munSetting 
         Caption         =   "设置"
      End
      Begin VB.Menu munExit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "FrmSetting"
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
'        Date : 2 0 1 2 / 0 4 / 1 2
' Description :
'     History :
'*****************************************************************
Option Explicit

Private Sub chkAutoAddUserCode_Click()

    If chkAutoAddUserCode.value = Checked Then
        optPrefix.Enabled = True
        optSuffix.Enabled = True
        comUserCode.Enabled = True
        picUserCode.Enabled = True
        picUserCode.MousePointer = 3
        FrmMain.chkAddUserCode.value = Checked
    Else
        optPrefix.Enabled = False
        optSuffix.Enabled = False
        comUserCode.Enabled = False
        picUserCode.Enabled = False
        picUserCode.MousePointer = 0
        FrmMain.chkAddUserCode.value = Unchecked

    End If

End Sub

Private Sub chkAutoUseDomain_Click()

    If chkAutoUseDomain.value = Checked Then
        chkDomainSuffix.Enabled = True
        chkAutoCheckClipboard.Enabled = True
    Else
        chkDomainSuffix.Enabled = False
        chkAutoCheckClipboard.Enabled = False

    End If

End Sub

Private Sub chkUseMouseHook_Click()

    If (Not isUseMouseHook) And chkUseMouseHook.value = Checked Then
        If MsgBox("此操作与系统密切相关, 可能会引起安全软件警告. 确定要使用该功能吗?" & vbCrLf & vbCrLf & "注: 开启该功能后, 为了能准确模拟按键输入, 花密主面板会在鼠标左键点击后自动隐藏.", vbQuestion + vbYesNo + vbDefaultButton2 + vbSystemModal) = vbYes Then
            chkUseMouseHook.value = Checked
        Else
            chkUseMouseHook.value = Unchecked

        End If

    End If

End Sub

Private Sub cmdCancel_Click()
    Call LoadSetting
    Call FrmMain.ShowResult
    Unload Me

End Sub

Private Sub cmdConfirm_Click()
    Call SaveSetting
    Call FrmMain.ShowResult
    Unload Me

End Sub

Private Sub cmdDefault_Click()

    Dim strSettingPath As String, isExist As Boolean

    strSettingPath = App.Path + "\Config.ini"
    isExist = False
    
    If Dir(strSettingPath, vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
        isExist = True
        Name strSettingPath As strSettingPath + ".bak"

    End If

    Call LoadSetting
    Call FrmMain.ShowResult
    
    If isExist Then
        Name strSettingPath + ".bak" As strSettingPath
        
    End If

End Sub

Private Sub comUserCode_Change()

    If Len(comUserCode.Text) > 0 Then
        picUserCode.Visible = False
    Else
        picUserCode.Visible = True

    End If

    If Not isUserCodeLoading Then cbBox_Change comUserCode

End Sub

Private Sub comUserCode_Click()

    If Len(comUserCode.Text) > 0 Then
        picUserCode.Visible = False
    Else
        picUserCode.Visible = True

    End If

End Sub

Private Sub comUserCode_KeyDown(KeyCode As Integer, Shift As Integer)
    cbBox_KeyDown comUserCode, KeyCode, Shift

    If Len(comUserCode.Text) > 0 Then
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

Private Sub Form_Load()
    picUserCode.Print "请输入附加扰码"
    Call LoadSetting

End Sub

Private Sub munExit_Click()
    Unload Me
    Unload FrmMain

End Sub

Private Sub munSetting_Click()
    Me.Show

End Sub

Private Sub munShow_Click()

    If Me.Visible Then
        MsgBox "请先关闭系统设置窗口, 再尝试显示花密主面板!", vbCritical + vbOKOnly + vbSystemModal
    Else
        Call FrmMain.MidScreenShow

    End If

End Sub

Private Sub picUserCode_Click()

    If comUserCode.Enabled Then comUserCode.SetFocus

End Sub
