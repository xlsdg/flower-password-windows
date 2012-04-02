Attribute VB_Name = "modEdit"
Option Explicit

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function GetWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal wCmd As Long) As Long

Private Declare Function GetClassName _
                Lib "user32" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Private Const GW_CHILD = 5

Private Const GW_HWNDNEXT = 2

Private Const WM_GETTEXT = &HD

Private Const WM_GETTEXTLENGTH = &HE

' Return information about this window.
Public Function WindowInfo(ByVal window_hwnd As Long, _
                           ByVal class_name As String) As String

    Static txt As String

    Dim buf    As String

    Dim buflen As Long

    ' Get the class name.
    buflen = 256
    buf = Space$(buflen - 1)
    buflen = GetClassName(window_hwnd, buf, buflen)
    buf = Left$(buf, buflen)

    If LCase$(buf) = class_name Then
        ' Associated text.
        txt = WindowText(window_hwnd)
        Exit Function

    End If

    Dim child_hwnd As Long

    Dim children   As Collection

    ' Make a list of the child windows.
    Set children = New Collection
    child_hwnd = GetWindow(window_hwnd, GW_CHILD)

    Do While child_hwnd <> 0
        children.Add child_hwnd
        child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
    Loop

    Dim I As Integer

    ' Get information on the child windows.
    For I = 1 To children.Count
        WindowInfo children(I), class_name
    Next I

    WindowInfo = txt

End Function

' Return the text associated with the window.
Public Function WindowText(ByVal window_hwnd As Long) As String

    Dim txtlen As Long

    Dim txt    As String

    WindowText = vbNullString

    If window_hwnd = 0 Then Exit Function
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)

    If txtlen = 0 Then Exit Function
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
    WindowText = Left$(txt, txtlen)

End Function
