Attribute VB_Name = "modData"
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

Public strDomains   As String

Public strPassWords As String

Public Sub LoadDatabase()
    Call LoadDomains
    Call LoadPasswords

End Sub

Private Function LoadDomains() As Boolean
    Dim bytData() As Byte

    bytData = LoadResData("DOMAINS", "DATA")
    strDomains = StrConv(bytData, vbUnicode)

    If Len(strDomains) > 0 Then
        LoadDomains = True
    Else
        LoadDomains = False

    End If

End Function

Private Function LoadPasswords() As Boolean
    Dim bytData() As Byte

    bytData = LoadResData("PASSWORDS", "DATA")
    strPassWords = StrConv(bytData, vbUnicode)

    If Len(strPassWords) > 0 Then
        LoadPasswords = True
    Else
        LoadPasswords = False

    End If

End Function

