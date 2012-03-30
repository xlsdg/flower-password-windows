Attribute VB_Name = "modOpera"
Option Explicit

Private Declare Function GetClassName _
                Lib "user32.dll" _
                Alias "GetClassNameA" (ByVal hWnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Public Function GetOperaDomainName(ByVal hWnd As Long) As String
    GetOperaDomainName = isClipboardAsUrl() 'GetWebsiteName(vbNullString)

End Function

Public Function isOpera(ByVal WinWnd As Long) As Boolean

    If WinWnd > 0 Then

        Dim RetVal As Long, lpClassName As String

        lpClassName = Space$(256)
        RetVal = GetClassName(WinWnd, lpClassName, 256)

        If InStr(Left$(lpClassName, RetVal), "OperaWindowClass") > 0 Then
            isOpera = True
        Else
            isOpera = False

        End If

    Else
        isOpera = False

    End If

End Function
