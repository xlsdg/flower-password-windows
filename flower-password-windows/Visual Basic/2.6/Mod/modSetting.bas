Attribute VB_Name = "modSetting"
Option Explicit

Public isAutoCopy        As Boolean

Public isAutoUseDomain   As Boolean

Public isDomainSuffix    As Boolean

Public isAutoAddUserCode As Boolean

Public isPrefix          As Boolean

Public isSuffix          As Boolean

Public isSetting         As Boolean

Public Sub LoadSetting()

    Dim strSettingPath As String

    strSettingPath = App.Path + "\Config.ini"

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

    If ReadIni("Setting", "DomainSuffix", strSettingPath) = "0" Then
        isDomainSuffix = False
        FrmMain.chkDomainSuffix.value = Unchecked
    Else
        isDomainSuffix = True
        FrmMain.chkDomainSuffix.value = Checked

    End If

    If ReadIni("Setting", "AutoAddCode", strSettingPath) = "0" Then
        isAutoAddUserCode = False
        FrmMain.chkAutoAddUserCode.value = Unchecked
        FrmMain.chkAddUserCode.value = Unchecked
    Else
        isAutoAddUserCode = True
        FrmMain.chkAutoAddUserCode.value = Checked
        FrmMain.chkAddUserCode.value = Checked

    End If

    If ReadIni("Setting", "Prefix", strSettingPath) = "0" Then
        isPrefix = False
        FrmMain.optPrefix.value = False
    Else
        isPrefix = True
        FrmMain.optPrefix.value = True

    End If

    If isPrefix Then
        isSuffix = False
        FrmMain.optSuffix.value = False
    Else
        isSuffix = True
        FrmMain.optSuffix.value = True

    End If

    FrmMain.txtUserCode.Text = ReadIni("Setting", "UserCode", strSettingPath)

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
        WriteIni "Setting", "UserCode", FrmMain.txtUserCode.Text, strSettingPath
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

End Sub
