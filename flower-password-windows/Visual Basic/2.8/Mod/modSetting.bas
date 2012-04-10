Attribute VB_Name = "modSetting"
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

Public isAutoMini        As Boolean

Public isAutoCopy        As Boolean

Public isAutoUseDomain   As Boolean

Public isDomainSuffix    As Boolean

Public isAutoAddUserCode As Boolean

Public isPrefix          As Boolean

Public isSuffix          As Boolean

Public isSetting         As Boolean

Public isLoading         As Boolean

Public isExit            As Boolean

Public Function AddStrToCombox(cbBox As ComboBox, ByVal strKey As String) As Boolean

    If Len(strKey) > 0 Then
        If CheckCombox(cbBox, strKey) Then
            AddStrToCombox = False
        Else
            cbBox.AddItem strKey
            AddStrToCombox = True

        End If

    End If

End Function

Public Function CheckCombox(cbBox As ComboBox, ByVal strKey As String) As Boolean

    Dim i As Long

    CheckCombox = False

    For i = 0 To cbBox.ListCount - 1

        If cbBox.List(i) = strKey Then
            CheckCombox = True
            Exit For

        End If

    Next

End Function

Public Sub LoadSetting()
    isExit = False
    Call LoadData

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    If ReadIni("Setting", "AutoMini", strSettingPath) = "1" Then
        isAutoMini = True
    Else
        isAutoMini = False

    End If

    If ReadIni("Setting", "AutoCopy", strSettingPath) = "0" Then
        isAutoCopy = False
        FrmMain.chkAutoCopy.value = Unchecked
    Else
        isAutoCopy = True
        FrmMain.chkAutoCopy.value = Checked

    End If

    If ReadIni("Setting", "AutoUseDomain", strSettingPath) = "0" Then
        isAutoUseDomain = False
        FrmMain.chkAutoUseDomain.value = Unchecked
    Else
        isAutoUseDomain = True
        FrmMain.chkAutoUseDomain.value = Checked

    End If

    If ReadIni("Setting", "DomainSuffix", strSettingPath) = "1" Then
        isDomainSuffix = True
        FrmMain.chkDomainSuffix.value = Checked
    Else
        isDomainSuffix = False
        FrmMain.chkDomainSuffix.value = Unchecked

    End If

    If ReadIni("Setting", "AutoAddCode", strSettingPath) = "1" Then
        isAutoAddUserCode = True
        FrmMain.chkAutoAddUserCode.value = Checked
        FrmMain.chkAddUserCode.value = Checked
    Else
        isAutoAddUserCode = False
        FrmMain.chkAutoAddUserCode.value = Unchecked
        FrmMain.chkAddUserCode.value = Unchecked

    End If

    If ReadIni("Setting", "Prefix", strSettingPath) = "1" Then
        isPrefix = True
        FrmMain.optPrefix.value = True
    Else
        isPrefix = False
        FrmMain.optPrefix.value = False

    End If

    If isPrefix Then
        isSuffix = False
        FrmMain.optSuffix.value = False
    Else
        isSuffix = True
        FrmMain.optSuffix.value = True

    End If

    isLoading = True
    Call LoadUserCode
    isLoading = False

End Sub

Public Sub SaveSetting()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    If FrmMain.chkAutoCopy.value = Checked Then
        isAutoCopy = True
        WriteIni "Setting", "AutoCopy", "1", strSettingPath
    Else
        isAutoCopy = False
        WriteIni "Setting", "AutoCopy", "0", strSettingPath

    End If

    If FrmMain.chkAutoUseDomain.value = Checked Then
        isAutoUseDomain = True
        WriteIni "Setting", "AutoUseDomain", "1", strSettingPath
    Else
        isAutoUseDomain = False
        WriteIni "Setting", "AutoUseDomain", "0", strSettingPath

    End If

    If FrmMain.chkDomainSuffix.value = Checked Then
        isDomainSuffix = True
        WriteIni "Setting", "DomainSuffix", "1", strSettingPath
    Else
        isDomainSuffix = False
        WriteIni "Setting", "DomainSuffix", "0", strSettingPath

    End If

    If FrmMain.chkAutoAddUserCode.value = Checked Then
        FrmMain.chkAddUserCode.value = Checked
        isAutoAddUserCode = True
        WriteIni "Setting", "AutoAddCode", "1", strSettingPath
    Else
        FrmMain.chkAddUserCode.value = Unchecked
        isAutoAddUserCode = False
        WriteIni "Setting", "AutoAddCode", "0", strSettingPath

    End If

    If FrmMain.optPrefix.value = True Then
        isPrefix = True
        WriteIni "Setting", "Prefix", "1", strSettingPath
        isSuffix = False
        WriteIni "Setting", "Suffix", "0", strSettingPath
    ElseIf FrmMain.optSuffix.value = True Then
        isPrefix = False
        WriteIni "Setting", "Prefix", "0", strSettingPath
        isSuffix = True
        WriteIni "Setting", "Suffix", "1", strSettingPath

    End If

    AddStrToCombox FrmMain.comUserCode, FrmMain.comUserCode.Text
    Call SaveUserCode

End Sub

Private Sub LoadUserCode()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    Dim strUserCode As String

    strUserCode = ReadIni("Setting", "UserCode", strSettingPath)

    If Len(strUserCode) > 0 Then

        Dim strCode() As String

        strCode = Split(strUserCode, Chr$(1), -1, vbBinaryCompare)

        Dim index As Long

        For index = LBound(strCode) To UBound(strCode)

            If Len(strCode(index)) > 0 Then
                FrmMain.comUserCode.AddItem strCode(index)

            End If

        Next
        strUserCode = ReadIni("Setting", "LastUserCode", strSettingPath)

        If Len(strUserCode) > 0 Then
            index = CLng(strUserCode)

            If 0 <= index And index <= FrmMain.comUserCode.ListCount - 1 Then
                FrmMain.comUserCode.Text = FrmMain.comUserCode.List(index)

            End If

        End If

    End If

End Sub

Private Sub SaveUserCode()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    Dim index As Long, strCode As String

    strCode = ""

    For index = 0 To FrmMain.comUserCode.ListCount - 1

        If index + 1 <> FrmMain.comUserCode.ListCount Then
            strCode = strCode + FrmMain.comUserCode.List(index) + Chr$(1)
        Else
            strCode = strCode + FrmMain.comUserCode.List(index)

        End If

    Next
    WriteIni "Setting", "UserCode", strCode, strSettingPath

    If FrmMain.comUserCode.ListIndex < 0 Then
        WriteIni "Setting", "LastUserCode", FrmMain.comUserCode.ListCount - 1, strSettingPath
    Else
        WriteIni "Setting", "LastUserCode", FrmMain.comUserCode.ListIndex, strSettingPath

    End If

End Sub
