Attribute VB_Name = "modTray"
Option Explicit

Private Declare Function Shell_NotifyIcon _
                Lib "shell32.dll" _
                Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                           lpData As NOTIFYICONDATA) As Long

'---------- dwMessage可以是以下NIM_ADD、NIM_DELETE、NIM_MODIFY 标识符之一----------
Private Const NIM_ADD = &H0    '在任务栏中增加一个图标

Private Const NIM_DELETE = &H2    '删除任务栏中的一个图标

Private Const NIM_MODIFY = &H1    '修改任务栏中个图标信息

Private Const NIM_SETFOCUS = &H3

'Private Const NIM_SETVERSION = &H4
Private Const NIF_MESSAGE = &H1    'NOTIFYICONDATA结构中uFlags的控制信息

Private Const NIF_ICON = &H2

Private Const NIF_TIP = &H4

'
Private Const NIF_STATE = &H8

Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1

Private Const NIS_SHAREDICON = &H2

Private Const WM_MOUSEMOVE = &H200

'Private Const WM_LBUTTONDOWN = &H201
'Private Const WM_LBUTTONUP = &H202
'Private Const WM_LBUTTONDBLCLK = &H203
'Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

'Private Const WM_RBUTTONDBLCLK = &H206
''Private Const WM_MBUTTONDOWN = &H207
'Private Const WM_MBUTTONUP = &H208
'Private Const WM_MBUTTONDBLCLK = &H209
'Private Const NOTIFYICON_VERSION = 3    '风格
'Private Const NOTIFYICON_OLDVERSION = 0    'Win95 任务栏样式
'系统托盘类型
Private Type NOTIFYICONDATA

    cbSize As Long    '该数据结构的大小
    hwnd As Long    '处理任务栏中图标的窗口句柄
    uID As Long    '定义的任务栏中图标的标识
    uFlags As Long    '任务栏图标功能控制，可以是以下值的组合（一般全包括）
    '   NIF_MESSAGE 表示发送控制消息；
    '   NIF_ICON表示显示控制栏中的图标；
    '   NIF_TIP表示任务栏中的图标有动态提示。
    uCallbackMessage As Long    '任务栏图标通过它与用户程序交换消息，处理该消息的窗口由hWnd决定
    hIcon As Long    '任务栏中的图标的控制句柄
    szTip As String * 128    '图标的提示信息。若要产生气泡提示信息，则一定要128才性，为64则无法生成气泡，其它功能都正常，原因不明
    '气泡提示信息部分
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256    '气泡提示内容
    uTimeout As Long    '气泡提示显示时间
    szInfoTitle As String * 64    '气泡提示标题
    dwInfoFlags As Long    '气泡提示类型，见 NIIF_*** 部分

End Type

Public Enum ico '气泡提示类型

    NIIF_NONE = &H0     'w无图标 = &H0      '
    NIIF_INFO = &H1     'x信息图标 = &H1    '
    NIIF_WARNING = &H2  'j警告图标 = &H2    '
    NIIF_ERROR = &H3    'z错误图标 = &H3    '
    NIIF_GUID = &H4     't托盘图标 = &H4    '

End Enum

Private IconData As NOTIFYICONDATA

Public Sub AddToTray(ByVal frm As Form, ByVal Tip As String, Optional ByVal TrayIco = 0)

    '生成系统托盘图标
    With IconData
        .cbSize = Len(IconData)
        .hwnd = frm.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE    '响应鼠标事件 'WM_LBUTTONDOWN

        If TrayIco = 0 Then
            .hIcon = frm.Icon    '默认为窗口的图标
        Else
            .hIcon = TrayIco

        End If

        .szTip = Tip & vbNullChar

    End With

    Shell_NotifyIcon NIM_ADD, IconData    '增加托盘图标

End Sub

Public Sub SetTrayMsgbox(ByVal MsgInfo As String, _
                         ByVal MsgFlags As Integer, _
                         ByVal MsgTitle As String, _
                         Optional ByVal TrayIco = 0)

    '    "系统托盘气泡提示文字不得超过128个字符！"
    With IconData
        .szInfoTitle = MsgTitle & Chr(0)
        .szInfo = MsgInfo & Chr(0)
        .dwInfoFlags = MsgFlags

        If TrayIco <> 0 Then
            .hIcon = TrayIco    '更换托盘图标

        End If

    End With

    Shell_NotifyIcon NIM_MODIFY, IconData    '修改托盘图标及相关信息

End Sub

Public Sub RemoveFromTray()
    Shell_NotifyIcon NIM_DELETE, IconData    '卸载托盘图标

End Sub
