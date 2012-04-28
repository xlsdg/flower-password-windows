Attribute VB_Name = "modFlowerPassword"
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

Public Function getFlowerPassword(ByVal strPassword As String, _
                                  ByVal strKey As String, _
                                  Optional ByVal PasswordBit As Integer = 16) As String

    If Len(strPassword) > 0 And Len(strKey) > 0 Then

        Dim cMD5     As New clsMD5

        Dim md5one   As String

        Dim md5two   As String

        Dim md5three As String

        md5one = LCase$(cMD5.Hmac_MD5(strPassword, strKey))
        md5two = LCase$(cMD5.Hmac_MD5(md5one, "snow"))
        md5three = LCase$(cMD5.Hmac_MD5(LCase$(md5one), "kise"))

        Dim code32 As String

        code32 = vbNullString

        Dim i As Integer

        For i = 1 To Len(md5two)

            If Not IsNumeric(Mid$(md5two, i, 1)) Then

                Dim str As String

                str = "sunlovesnow1990090127xykab"

                If InStr(1, str, Mid$(md5three, i, 1), vbBinaryCompare) > 0 Then
                    code32 = code32 + UCase$(Mid$(md5two, i, 1))
                Else
                    code32 = code32 + Mid$(md5two, i, 1)

                End If

            Else
                code32 = code32 + Mid$(md5two, i, 1)

            End If

        Next

        Dim code1 As String

        code1 = Left$(code32, 1)

        Dim code16 As String

        If Not IsNumeric(code1) Then
            code16 = Left$(code32, 16)
        Else
            code16 = "K" + Mid$(code32, 2, 15)

        End If

        If PasswordBit = 16 Then
            getFlowerPassword = code16
        Else
            getFlowerPassword = code32

        End If

        Set cMD5 = Nothing

    End If

End Function
