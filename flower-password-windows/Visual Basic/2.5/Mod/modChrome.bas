Attribute VB_Name = "modChrome"
Option Explicit

Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Private Declare Function GetClassName _
                Lib "user32.dll" _
                Alias "GetClassNameA" (ByVal hWnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Public Function GetChromeDomainName(ByVal hWnd As Long) As String

    Dim strUrl As String

    strUrl = WindowInfo(GetForegroundWindow, LCase$("Chrome_OmniboxView"))

    If Len(strUrl) > 0 Then
        GetChromeDomainName = GetWebsiteName(strUrl)
    Else
        GetChromeDomainName = vbNullString

    End If

End Function

Public Function isChrome(ByVal WinWnd As Long) As Boolean

    If WinWnd > 0 Then

        Dim RetVal As Long, lpClassName As String

        lpClassName = Space$(256)
        RetVal = GetClassName(WinWnd, lpClassName, 256)

        If InStr(Left$(lpClassName, RetVal), "Chrome_RenderWidgetHostHWND") > 0 Then
            isChrome = True
        Else
            isChrome = False

        End If

    Else
        isChrome = False

    End If

End Function
