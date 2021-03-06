Attribute VB_Name = "modFirefox"
'*****************************************************************
' Copyright (c) 2011-2012 FlowerPassword.com All rights reserved.
'      Author : xLsDg @ Xiao Lu Software Development Group
'        Blog : http://hi.baidu.com/xlsdg
'          QQ : 4 4 7 4 0 5 7 4 0
'     Version : 1 . 0 . 0 . 0
'        Date : 2 0 1 2 / 0 4 / 0 7
' Description :
'     History :
'*****************************************************************
Option Explicit

Private Declare Function GetClassName _
                Lib "user32.dll" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Public Function GetFirefoxDomainName(ByVal hwnd As Long) As String
    GetFirefoxDomainName = isClipboardAsUrl() 'GetWebsiteName(vbNullString)

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
