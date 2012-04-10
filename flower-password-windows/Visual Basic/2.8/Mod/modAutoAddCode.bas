Attribute VB_Name = "modAutoAddCode"
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

Private blnAuto As Boolean

Public Function GetKeyFromStr(ByVal strUser As String) As String

    Dim strCode As String, lenCode As Long

    strCode = FrmMain.comUserCode.Text
    lenCode = Len(strCode)

    If Left$(strUser, lenCode) = strCode Then
        strUser = Right$(strUser, Len(strUser) - lenCode)

    End If

    If Right$(strUser, lenCode) = strCode Then
        strUser = Left$(strUser, Len(strUser) - lenCode)

    End If

    GetKeyFromStr = strUser

End Function

Public Sub KeyBox_AutoComplete(cbBox As ComboBox)

    If blnAuto Then
        blnAuto = False
    Else

        With cbBox

            Dim lenBoxText As Long

            lenBoxText = Len(.Text)

            If lenBoxText > 0 Then

                Dim strCode As String, lenCode As Long, lenKey As Long

                strCode = FrmMain.comUserCode.Text
                lenCode = Len(strCode)
                lenKey = Len(GetKeyFromStr(.Text))

                If .Text <> strCode Then

                    Dim strPartial As String, iStart As Long

                    iStart = .SelStart

                    If isAutoAddUserCode Then
                        If isPrefix Then
                            If Right$(.Text, lenCode) = strCode Then
                                If iStart > lenCode + lenKey Then
                                    iStart = lenCode + lenKey

                                End If

                                blnAuto = True
                                .Text = Left$(.Text, lenBoxText - lenCode)
                                blnAuto = False

                            End If

                            If Left$(.Text, lenCode) <> strCode Then
                                blnAuto = True
                                .Text = strCode + .Text
                                blnAuto = False
                                iStart = lenCode + iStart
                            Else

                                If iStart < lenCode Then
                                    iStart = lenCode

                                End If

                            End If

                        Else

                            If Left$(.Text, lenCode) = strCode Then
                                If iStart < lenCode Then
                                    iStart = 0

                                End If

                                blnAuto = True
                                .Text = Right$(.Text, lenBoxText - lenCode)
                                blnAuto = False

                            End If

                            If Right$(.Text, lenCode) <> strCode Then
                                blnAuto = True
                                .Text = .Text + strCode
                                blnAuto = False
                            Else

                                If iStart > lenKey Then
                                    iStart = lenKey

                                End If

                            End If

                        End If

                        strPartial = GetKeyFromStr(.Text)
                    Else

                        If Left$(.Text, lenCode) = strCode Then
                            blnAuto = True
                            .Text = Right$(.Text, lenBoxText - lenCode)
                            blnAuto = False

                        End If

                        If Right$(.Text, lenCode) = strCode Then
                            blnAuto = True
                            .Text = Left$(.Text, lenBoxText - lenCode)
                            blnAuto = False

                        End If

                        strPartial = .Text

                    End If

                    blnAuto = True
                    .SelStart = iStart
                    blnAuto = False

                    Dim i As Long

                    For i = 0 To .ListCount - 1

                        Dim strTotal As String

                        strTotal = .List(i)

                        If (strTotal Like (strPartial & "*")) And (strTotal <> strPartial) Then

                            Dim J As Long

                            J = Len(strTotal) - Len(strPartial)

                            If J <> 0 Then
                                blnAuto = True
                                .SelText = Right$(strTotal, J)

                                If isAutoAddUserCode Then
                                    If isPrefix Then
                                        .SelStart = lenCode + Len(strPartial)
                                    Else
                                        .SelStart = Len(strPartial)

                                    End If

                                Else
                                    .SelStart = Len(strPartial)

                                End If

                                .SelLength = J
                                blnAuto = False
                                Exit For

                            End If

                        End If

                    Next
                Else
                    blnAuto = True
                    .Text = ""
                    blnAuto = False

                End If

            End If

        End With

    End If

End Sub

Public Sub KeyBox_KeyDown(cbBox As ComboBox, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If isAutoAddUserCode Then

            Dim strKey As String, lenKey As Long

            strKey = GetKeyFromStr(cbBox.Text)
            lenKey = Len(strKey)

            If isPrefix Then
                If cbBox.SelStart <= Len(FrmMain.comUserCode.Text) Then
                    blnAuto = True
                    cbBox.SelStart = Len(cbBox.Text)
                    blnAuto = False

                End If

            Else

                If cbBox.SelStart > lenKey Then
                    blnAuto = True
                    cbBox.SelStart = lenKey
                    blnAuto = False

                End If

            End If

        End If

        blnAuto = True
        cbBox.SelText = ""
        blnAuto = False

    End If

End Sub
