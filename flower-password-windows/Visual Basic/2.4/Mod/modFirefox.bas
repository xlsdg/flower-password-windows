Attribute VB_Name = "modFirefox"
Option Explicit

Private Declare Function GetClassName _
                Lib "user32.dll" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Public Function GetFirefoxDomainName(ByVal hwnd As Long) As String
    GetFirefoxDomainName = "" 'GetWebsiteName(vbNullString)

End Function

Public Function isFirefox(ByVal WinWnd As Long) As Boolean

    If WinWnd > 0 Then

        Dim RetVal As Long, lpClassName As String

        lpClassName = Space$(256)
        RetVal = GetClassName(WinWnd, lpClassName, 256)

        If InStr(Left$(lpClassName, RetVal), "MozillaWindowClass") > 0 Then
            isFirefox = True
        Else
            isFirefox = False

        End If

    Else
        isFirefox = False

    End If

End Function
