Attribute VB_Name = "modMaxthon"
Option Explicit

Private Declare Function GetClassName _
                Lib "user32.dll" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Public Function GetMaxthonDomainName(ByVal hwnd As Long) As String
    GetMaxthonDomainName = isClipboardAsUrl() 'GetWebsiteName(vbNullString)

End Function

Public Function isMaxthon(ByVal WinWnd As Long) As Boolean

    If WinWnd > 0 Then

        Dim RetVal As Long, lpClassName As String

        lpClassName = Space$(256)
        RetVal = GetClassName(WinWnd, lpClassName, 256)

        If InStr(Left$(lpClassName, RetVal), "Maxthon3Cls_WebViewHost") > 0 Then
            isMaxthon = True
        Else
            isMaxthon = False

        End If

    Else
        isMaxthon = False

    End If

End Function
