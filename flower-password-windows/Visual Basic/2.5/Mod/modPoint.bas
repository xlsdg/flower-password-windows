Attribute VB_Name = "modPoint"
Option Explicit

Public Password_Hwnd As Long

Private Declare Function GetWindowRect _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  lpRect As Rect) As Long

Private Type Rect

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Private Type POINTAPI

    X As Long
    Y As Long

End Type

Private Declare Function WindowFromPoint _
                Lib "user32.dll" (ByVal xPoint As Long, _
                                  ByVal yPoint As Long) As Long

'获得拥有输入焦点的窗口的句柄
Private Declare Function GetFocus Lib "user32.dll" () As Long

'获得前台窗口的句柄。这里的“前台窗口”是指前台应用程序的活动窗口
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

'通常，系统内的每个线程都有自己的输入队列。本函数（既“连接线程输入函数”）允许线程和进程共享输入队列。连接了线程后，输入焦点、窗口激活、鼠标捕获、键盘状态以及输入队列状态都会进入共享状态
Private Declare Function AttachThreadInput _
                Lib "user32.dll" (ByVal idAttach As Long, _
                                  ByVal idAttachTo As Long, _
                                  ByVal fAttach As Long) As Long

'获取当前线程一个唯一的线程标识符
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long

'获取与指定窗口关联在一起的一个进程和线程标识符
Private Declare Function GetWindowThreadProcessId _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  lpdwProcessId As Long) As Long

'判断插入符的当前位置
Private Declare Function GetCaretPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

'判断窗口内以客户区坐标表示的一个点的屏幕坐标
Private Declare Function ClientToScreen _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  lpPoint As POINTAPI) As Long

Public Function GetDesktopWindowCaretPos(ByRef point As POINTAPI) As Long

    Dim foregroundWindowHandle As Long

    foregroundWindowHandle = GetForegroundWindow

    Dim remoteThreadId As Long

    remoteThreadId = GetWindowThreadProcessId(foregroundWindowHandle, 0)

    Dim currentThreadId As Long

    currentThreadId = GetCurrentThreadId()

    Dim result As Long

    result = AttachThreadInput(currentThreadId, remoteThreadId, True)

    If result <> 0 Then

        Dim focused As Long

        focused = GetFocus

        Dim ThisPoint As Long

        ThisPoint = GetCaretPos(point)
        ClientToScreen focused, point
        AttachThreadInput currentThreadId, remoteThreadId, False
        GetDesktopWindowCaretPos = focused
    Else
        GetDesktopWindowCaretPos = 0

    End If

End Function

Public Function GetDesktopWindowRect(ByRef rct As Rect, ByRef pos As POINTAPI) As Long
    GetCursorPos pos

    Dim WinHandle As Long

    WinHandle = WindowFromPoint(pos.X, pos.Y)

    Dim execute As Integer

    execute = GetWindowRect(WinHandle, rct)

    If execute = 0 Then
        GetDesktopWindowRect = 0
    Else
        GetDesktopWindowRect = WinHandle

    End If

End Function

Public Sub getLocation(ByRef point_x As Long, ByRef point_y As Long)

    Dim caretpos As POINTAPI, mousepos As POINTAPI, rects As Rect, caretpos_hWnd As Long, rect_hWnd As Long

    caretpos_hWnd = GetDesktopWindowCaretPos(caretpos)
    rect_hWnd = GetDesktopWindowRect(rects, mousepos)

    If caretpos_hWnd <> 0 Then
        Password_Hwnd = caretpos_hWnd

        If (mousepos.Y - caretpos.Y) < 50 Then
            point_x = caretpos.X '* Screen.TwipsPerPixelX
            point_y = (caretpos.Y + 17) '* Screen.TwipsPerPixelY
        Else
            point_x = mousepos.X '* Screen.TwipsPerPixelX
            point_y = (mousepos.Y + 10) '* Screen.TwipsPerPixelY

        End If

    ElseIf rect_hWnd <> 0 Then
        Password_Hwnd = rect_hWnd

        If (rects.Bottom - caretpos.Y) < 50 Then
            point_x = rects.Left '* Screen.TwipsPerPixelX
            point_y = rects.Bottom '* Screen.TwipsPerPixelY
        Else
            point_x = mousepos.X '* Screen.TwipsPerPixelX
            point_y = (mousepos.Y + 10) '* Screen.TwipsPerPixelY

        End If

    Else
        Password_Hwnd = 0
        point_x = mousepos.X '* Screen.TwipsPerPixelX
        point_y = (mousepos.Y + 10) '* Screen.TwipsPerPixelY

    End If

End Sub
