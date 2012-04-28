Attribute VB_Name = "modHSIMP"
'*****************************************************************
' Copyright (c) 2011-2012 FlowerPassword.com All rights reserved.
'      Author : xLsDg @ Xiao Lu Software Development Group
'        Blog : http://hi.baidu.com/xlsdg
'          QQ : 4 4 7 4 0 5 7 4 0
'     Version : 1 . 1 . 0 . 0
'        Date : 2 0 1 2 / 0 4 / 1 2
' Description :
'     History :
'*****************************************************************
Option Explicit

Public Const PASSWORD_WEAK   As Long = &H0

Public Const PASSWORD_NORMAL As Long = &H1

Public Const PASSWORD_STRONG As Long = &H2

Private Const PASSWORD_GOOD  As Long = -1

Public Function check_password_level(ByVal strPassword As String, _
                                     Optional ByRef strInfo As String) As Long

    Dim result(1 To 7) As Long, strTemp(1 To 7) As String

    '1.'Common Password'
    result(1) = check_common_password(strPassword, strTemp(1))
    '2.'Length'
    result(2) = check_length(strPassword, strTemp(2))
    '3.'Character Variety'
    result(3) = check_character_variety(strPassword, strTemp(3))
    '4.'Repeated Pattern'
    result(4) = check_repeated_pattern(strPassword, strTemp(4))
    '5.'Possibly a Word'
    result(5) = check_possibly_a_word(strPassword, strTemp(5))
    '6.'Possibly a Telephone Number / Date'
    result(6) = check_possibly_a_telephone_number_or_date(strPassword, strTemp(6))
    '7.'Possibly a Word and a Number'
    result(7) = check_possibly_a_word_and_a_number(strPassword, strTemp(7))

    Dim index As Long

    check_password_level = PASSWORD_STRONG

    For index = LBound(result) To UBound(result)

        If is_password_weak(check_password_level, result(index)) Then
            check_password_level = result(index)
            strInfo = strInfo & "■" & strTemp(index)

        End If

    Next

End Function

Private Function check_character_variety(ByVal strPassword As String, _
                                         Optional ByRef strInfo As String) As Long

    If password_match(strPassword, "^[a-zA-Z]+$") Then
        strInfo = "该密码只包含字母，加入数字和符号可提高强度"
        check_character_variety = PASSWORD_NORMAL
    ElseIf password_match(strPassword, "^[0-9]+$") Then
        strInfo = "该密码只包含数字，加入字母和符号可提高强度"
        check_character_variety = PASSWORD_NORMAL
    ElseIf password_match(strPassword, "^[A-Za-z0-9]+$") Then
        strInfo = "该密码只包含数字和字母，加入符号可提高强度"
        check_character_variety = PASSWORD_NORMAL
    Else
        check_character_variety = PASSWORD_GOOD

    End If

End Function

Private Function check_common_password(ByVal strPassword As String, _
                                       Optional ByRef strInfo As String) As Long

    Dim arrCommonPassword() As String, isFound As Boolean

    arrCommonPassword = Split(strPassWords, "|")
    isFound = False

    Dim index As Long

    For index = LBound(arrCommonPassword) To UBound(arrCommonPassword)

        If strPassword = arrCommonPassword(index) Then
            isFound = True
            Exit For

        End If

    Next

    If isFound Then
        strInfo = "该密码是常用密码之一，可被瞬间破解"
        check_common_password = PASSWORD_WEAK
    Else
        check_common_password = PASSWORD_GOOD

    End If

End Function

Private Function check_length(ByVal strPassword As String, _
                              Optional ByRef strInfo As String) As Long

    If Len(strPassword) < 5 Then
        strInfo = "该密码太短，请使用8位或以上的密码"
        check_length = PASSWORD_WEAK
    ElseIf Len(strPassword) < 8 Then
        strInfo = "该密码比较短，请使用8位或以上的密码"
        check_length = PASSWORD_NORMAL
    Else
        check_length = PASSWORD_GOOD

    End If

End Function

Private Function check_possibly_a_telephone_number_or_date(ByVal strPassword As String, _
                                                           Optional ByRef strInfo As String) As Long

    If password_match(strPassword, "^[\-\(\)\.\/\s0-9]+$") Then
        strInfo = "该密码可能是一个电话号码或一个日期，如果是的话，这将使其很容易被破解"
        check_possibly_a_telephone_number_or_date = PASSWORD_WEAK
    Else
        check_possibly_a_telephone_number_or_date = PASSWORD_GOOD

    End If

End Function

Private Function check_possibly_a_word(ByVal strPassword As String, _
                                       Optional ByRef strInfo As String) As Long

    If password_match(strPassword, "^[A-Za-z]+$") Then
        strInfo = "该密码可能是一个单词或一个名字，如果是的话，这将使其很容易被破解"
        check_possibly_a_word = PASSWORD_WEAK
    Else
        check_possibly_a_word = PASSWORD_GOOD

    End If

End Function

Private Function check_possibly_a_word_and_a_number(ByVal strPassword As String, _
                                                    Optional ByRef strInfo As String) As Long

    If password_match(strPassword, "^[a-zA-Z]+[0-9]+$") Or password_match(strPassword, "^[0-9]+[a-zA-Z]+$") Then
        strInfo = "该密码可能是一个单词和几个数字的组合，这是很常见的模式，因此可以被快速的破解"
        check_possibly_a_word_and_a_number = PASSWORD_WEAK
    Else
        check_possibly_a_word_and_a_number = PASSWORD_GOOD

    End If

End Function

Private Function check_repeated_pattern(ByVal strPassword As String, _
                                        Optional ByRef strInfo As String) As Long

    If password_match(strPassword, "(.+)\1{2,}", True, True) Then
        strInfo = "该密码包含重复的部分，这使其更容易被破解"
        check_repeated_pattern = PASSWORD_WEAK
    Else
        check_repeated_pattern = PASSWORD_GOOD

    End If

End Function

Private Function is_password_weak(ByVal Password_Level As Long, _
                                  ByVal Result_Level As Long) As Boolean

    If Result_Level <> PASSWORD_GOOD Then
        If Password_Level > Result_Level Then
            is_password_weak = True
        Else
            is_password_weak = False

        End If

    Else
        is_password_weak = False

    End If

End Function

Private Function password_match(ByVal strPassword As String, _
                                ByVal strPattern As String, _
                                Optional ByVal blnGlobal As Boolean = False, _
                                Optional ByVal blnIgnoreCase As Boolean = False, _
                                Optional ByVal binMutilLine As Boolean = False) As Boolean

    Dim objRegExp As RegExp

    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = strPattern
        .Global = blnGlobal
        .IgnoreCase = blnIgnoreCase
        .MultiLine = binMutilLine
        password_match = .Test(strPassword)

    End With

    Set objRegExp = Nothing

End Function
