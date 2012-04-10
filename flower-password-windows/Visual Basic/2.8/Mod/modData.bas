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

Public Sub LoadData()
    Call LoadDomains
    Call LoadPasswords

End Sub

Private Function LoadDomains() As Boolean

    Dim strPath As String

    strPath = App.Path + "\Domains.dat"
    ReleaseDataFromRes "DATA", "DOMAINS", strPath
    strDomains = ReadDataFromFile(strPath)
    Kill strPath

    If Len(strDomains) > 0 Then
        LoadDomains = True
    Else
        LoadDomains = False

    End If

End Function

Private Function LoadPasswords() As Boolean

    Dim strPath As String

    strPath = App.Path + "\Domains.dat"
    ReleaseDataFromRes "DATA", "PASSWORDS", strPath
    strPassWords = ReadDataFromFile(strPath)
    Kill strPath

    If Len(strPassWords) > 0 Then
        LoadPasswords = True
    Else
        LoadPasswords = False

    End If

End Function

Private Function ReadDataFromFile(ByVal strFilePath As String) As String

    If Len(strFilePath) > 0 Then
        If Dir$(strFilePath, vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then

            Dim bytData() As Byte

            Open strFilePath For Binary Access Read As #1
            ReDim bytData(1 To LOF(1)) As Byte
            Get #1, , bytData
            Close #1
            ReadDataFromFile = StrConv(bytData, vbUnicode)

        End If

    Else
        ReadDataFromFile = vbNullString

    End If

End Function

Private Function ReleaseDataFromRes(ByVal strType As String, _
                                    ByVal strID As String, _
                                    ByVal strFilePath As String) As Boolean

    If Len(strType) > 0 And Len(strID) > 0 And Len(strFilePath) > 0 Then
        If Dir$(strFilePath, vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
            Kill strFilePath

        End If

        Dim bytData() As Byte

        bytData = LoadResData(strID, strType)
        Open strFilePath For Binary Access Write As #1
        Put #1, , bytData
        Close #1
        ReleaseDataFromRes = True
    Else
        ReleaseDataFromRes = False

    End If

End Function

