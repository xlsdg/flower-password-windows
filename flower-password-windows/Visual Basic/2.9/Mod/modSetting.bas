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

Public isAlwaysOnTop     As Boolean

Public isCloseToHide     As Boolean

Public isAutoCopy        As Boolean

Public isAutoUseDomain   As Boolean

Public isDomainSuffix    As Boolean

Public isAutoCheck       As Boolean

Public isAutoAddUserCode As Boolean

Public isPrefix          As Boolean

Public isUseMouseHook    As Boolean

Public isUserCodeLoading As Boolean '正在加载附加扰码

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
    Call LoadDatabase '加载数据库

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    If ReadIni("Setting", "AutoMini", strSettingPath) = "1" Then
        isAutoMini = True
        FrmSetting.chkAutoMini.value = Checked
    Else
        isAutoMini = False
        FrmSetting.chkAutoMini.value = Unchecked

    End If

    If ReadIni("Setting", "AlwaysOnTop", strSettingPath) = "0" Then
        isAlwaysOnTop = False
        FrmSetting.chkAlwaysOnTop.value = Unchecked
    Else
        isAlwaysOnTop = True
        FrmSetting.chkAlwaysOnTop.value = Checked

    End If

    If ReadIni("Setting", "Transparent", strSettingPath) = "1" Then
        SetFrmTransparent FrmMain.hwnd
        FrmSetting.chkTransparent.value = Checked
    Else
        UnSetFrmTransparent FrmMain.hwnd
        FrmSetting.chkTransparent.value = Unchecked

    End If
    
    If ReadIni("Setting", "ShowCode", strSettingPath) = "0" Then
        FrmMain.lblCode16.Visible = False
        FrmSetting.chkShowCode.value = Unchecked
    Else
        FrmMain.lblCode16.Visible = True
        FrmSetting.chkShowCode.value = Checked

    End If
    
    If ReadIni("Setting", "CloseToHide", strSettingPath) = "0" Then
        isCloseToHide = False
        FrmSetting.OptCloseToExit.value = True
    Else
        isCloseToHide = True
        FrmSetting.OptClodeToHide.value = True

    End If

    If ReadIni("Setting", "AutoCopy", strSettingPath) = "0" Then
        isAutoCopy = False
        FrmSetting.chkAutoCopy.value = Unchecked
    Else
        isAutoCopy = True
        FrmSetting.chkAutoCopy.value = Checked

    End If

    If ReadIni("Setting", "AutoUseDomain", strSettingPath) = "0" Then
        isAutoUseDomain = False
        FrmSetting.chkAutoUseDomain.value = Unchecked
    Else
        isAutoUseDomain = True
        FrmSetting.chkAutoUseDomain.value = Checked

    End If

    If ReadIni("Setting", "DomainSuffix", strSettingPath) = "1" Then
        isDomainSuffix = True
        FrmSetting.chkDomainSuffix.value = Checked
    Else
        isDomainSuffix = False
        FrmSetting.chkDomainSuffix.value = Unchecked

    End If

    If ReadIni("Setting", "AutoCheckClipboard", strSettingPath) = "0" Then
        isAutoCheck = False
        FrmSetting.chkAutoCheckClipboard.value = Unchecked
    Else
        isAutoCheck = True
        FrmSetting.chkAutoCheckClipboard.value = Checked

    End If

    If ReadIni("Setting", "AutoAddCode", strSettingPath) = "1" Then
        isAutoAddUserCode = True
        FrmSetting.chkAutoAddUserCode.value = Checked
        FrmMain.chkAddUserCode.value = Checked
    Else
        isAutoAddUserCode = False
        FrmSetting.chkAutoAddUserCode.value = Unchecked
        FrmMain.chkAddUserCode.value = Unchecked

    End If

    If ReadIni("Setting", "Prefix", strSettingPath) = "1" Then
        isPrefix = True
        FrmSetting.optPrefix.value = True
        FrmSetting.optSuffix.value = False
    Else
        isPrefix = False
        FrmSetting.optPrefix.value = False
        FrmSetting.optSuffix.value = True

    End If

    If ReadIni("Setting", "MouseHook", strSettingPath) = "1" Then
        isUseMouseHook = True
        FrmSetting.chkUseMouseHook.value = Checked
    Else
        isUseMouseHook = False
        FrmSetting.chkUseMouseHook.value = Unchecked

    End If

    isUserCodeLoading = True
    Call LoadUserCode
    isUserCodeLoading = False

End Sub

Public Sub SaveSetting()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    If FrmSetting.chkAutoMini.value = Checked Then
        isAutoMini = True
        WriteIni "Setting", "AutoMini", "1", strSettingPath
    Else
        isAutoMini = False
        WriteIni "Setting", "AutoMini", "0", strSettingPath

    End If

    If FrmSetting.chkAlwaysOnTop.value = Checked Then
        isAlwaysOnTop = True
        WriteIni "Setting", "AlwaysOnTop", "1", strSettingPath
    Else
        isAlwaysOnTop = False
        WriteIni "Setting", "AlwaysOnTop", "0", strSettingPath

    End If

    If FrmSetting.chkTransparent.value = Checked Then
        SetFrmTransparent FrmMain.hwnd
        WriteIni "Setting", "Transparent", "1", strSettingPath
    Else
        UnSetFrmTransparent FrmMain.hwnd
        WriteIni "Setting", "Transparent", "0", strSettingPath

    End If

    If FrmSetting.chkShowCode.value = Checked Then
        FrmMain.lblCode16.Visible = True
        WriteIni "Setting", "ShowCode", "1", strSettingPath
    Else
        FrmMain.lblCode16.Visible = False
        WriteIni "Setting", "ShowCode", "0", strSettingPath

    End If

    If FrmSetting.OptClodeToHide.value = True Then
        isCloseToHide = True
        WriteIni "Setting", "CloseToHide", "1", strSettingPath
    Else
        isCloseToHide = False
        WriteIni "Setting", "CloseToHide", "0", strSettingPath

    End If

    If FrmSetting.chkAutoCopy.value = Checked Then
        isAutoCopy = True
        WriteIni "Setting", "AutoCopy", "1", strSettingPath
    Else
        isAutoCopy = False
        WriteIni "Setting", "AutoCopy", "0", strSettingPath

    End If

    If FrmSetting.chkAutoUseDomain.value = Checked Then
        isAutoUseDomain = True
        WriteIni "Setting", "AutoUseDomain", "1", strSettingPath
    Else
        isAutoUseDomain = False
        WriteIni "Setting", "AutoUseDomain", "0", strSettingPath

    End If

    If FrmSetting.chkDomainSuffix.value = Checked Then
        isDomainSuffix = True
        WriteIni "Setting", "DomainSuffix", "1", strSettingPath
    Else
        isDomainSuffix = False
        WriteIni "Setting", "DomainSuffix", "0", strSettingPath

    End If

    If FrmSetting.chkAutoCheckClipboard.value = Checked Then
        isAutoCheck = True
        WriteIni "Setting", "AutoCheckClipboard", "1", strSettingPath
    Else
        isAutoCheck = False
        WriteIni "Setting", "AutoCheckClipboard", "0", strSettingPath

    End If

    If FrmSetting.chkAutoAddUserCode.value = Checked Then
        FrmMain.chkAddUserCode.value = Checked
        isAutoAddUserCode = True
        WriteIni "Setting", "AutoAddCode", "1", strSettingPath
    Else
        FrmMain.chkAddUserCode.value = Unchecked
        isAutoAddUserCode = False
        WriteIni "Setting", "AutoAddCode", "0", strSettingPath

    End If

    If FrmSetting.optPrefix.value = True Then
        isPrefix = True
        WriteIni "Setting", "Prefix", "1", strSettingPath
    Else
        isPrefix = False
        WriteIni "Setting", "Prefix", "0", strSettingPath

    End If

    If FrmSetting.chkUseMouseHook.value = Checked Then
        isUseMouseHook = True
        WriteIni "Setting", "MouseHook", "1", strSettingPath
    Else
        isUseMouseHook = False
        WriteIni "Setting", "MouseHook", "0", strSettingPath

    End If
        
    AddStrToCombox FrmSetting.comUserCode, FrmSetting.comUserCode.Text
    Call SaveUserCode

End Sub

Private Sub LoadUserCode()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    Dim strUserCode As String

    strUserCode = ReadIni("Setting", "UserCode", strSettingPath)

    If Len(strUserCode) > 0 Then
        FrmSetting.comUserCode.Clear

        Dim strCode() As String

        strCode = Split(strUserCode, Chr$(1), -1, vbBinaryCompare)

        Dim index As Long

        For index = LBound(strCode) To UBound(strCode)

            If Len(strCode(index)) > 0 Then
                FrmSetting.comUserCode.AddItem strCode(index)

            End If

        Next
        strUserCode = ReadIni("Setting", "LastUserCode", strSettingPath)

        If Len(strUserCode) > 0 Then
            index = CLng(strUserCode)

            If 0 <= index And index <= FrmSetting.comUserCode.ListCount - 1 Then
                FrmSetting.comUserCode.Text = FrmSetting.comUserCode.List(index)

            End If

        End If

    Else
        FrmSetting.comUserCode.Text = ""

    End If

End Sub

Private Sub SaveUserCode()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

    Dim index As Long, strCode As String, lstCount As Long

    strCode = vbNullString
    lstCount = FrmSetting.comUserCode.ListCount - 1

    For index = 0 To lstCount

        If index <> lstCount Then
            strCode = strCode + FrmSetting.comUserCode.List(index) + Chr$(1)
        Else
            strCode = strCode + FrmSetting.comUserCode.List(index)

        End If

    Next
    WriteIni "Setting", "UserCode", strCode, strSettingPath

    If FrmSetting.comUserCode.ListIndex < 0 Then
        WriteIni "Setting", "LastUserCode", FrmSetting.comUserCode.ListCount - 1, strSettingPath
    Else
        WriteIni "Setting", "LastUserCode", FrmSetting.comUserCode.ListIndex, strSettingPath

    End If

End Sub
