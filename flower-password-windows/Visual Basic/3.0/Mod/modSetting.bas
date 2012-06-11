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

Private Const MOD_ALT = &H1

Private Const MOD_CONTROL = &H2

Private Const MOD_SHIFT = &H4

Private Const MOD_WIN = &H8

'====================================PART0
Public isAutoMini        As Boolean

Public isShowHelp        As Boolean

Public isAlwaysOnTop     As Boolean

Public HotKeyValue       As Long

Public KeyValue          As Integer

'====================================PART1
Public isShowCode        As Boolean

Public isShowPassword    As Boolean

Public isCloseToHide     As Boolean

'====================================PART3
Public isAutoAddUserCode As Boolean

Public isPrefix          As Boolean

Public isSuffix          As Boolean

Public isUserCodeLoading As Boolean '正在加载附加扰码

'====================================PART4
Public isAutoCopy        As Boolean

Public isAutoUseDomain   As Boolean

Public isAutoCheck       As Boolean

Public isDomainSuffix    As Boolean

Public isProtect         As Boolean

Public isUseMouseHook    As Boolean

'====================================PART6
Private PasswordLength   As Integer

Public isDiyWordLength   As Boolean

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

Public Function calcHotKeyValue(ByVal blnShift As Boolean, _
                                ByVal blnCtrl As Boolean, _
                                ByVal blnWin As Boolean, _
                                ByVal blnAlt As Boolean) As Long

    Dim lngKey As Long

    lngKey = 0

    If blnShift Then
        lngKey = lngKey + MOD_SHIFT

    End If

    If blnCtrl Then
        lngKey = lngKey + MOD_CONTROL

    End If

    If blnWin Then
        lngKey = lngKey + MOD_WIN

    End If

    If blnAlt Then
        lngKey = lngKey + MOD_ALT

    End If

    If lngKey <> 0 Then
        calcHotKeyValue = lngKey
    Else
        calcHotKeyValue = MOD_WIN

    End If

End Function

Public Function calcPasswordLength(ByVal strKey As String) As Integer

    With FrmSetting

        Dim Index As Long, maxCount As Long, isFound As Boolean

        maxCount = .lstDiyKey.ListCount - 1
        isFound = False

        For Index = 0 To maxCount

            Dim strCode() As String

            strCode = Split(.lstDiyKey.List(Index), vbTab, -1, vbBinaryCompare)

            If strCode(0) = strKey Then
                isFound = True
                calcPasswordLength = CInt(strCode(1))
                Exit For

            End If

        Next

        If Not isFound Then
            If isDiyWordLength Then
                calcPasswordLength = PasswordLength
            Else
                calcPasswordLength = 16
        
            End If

        End If

    End With

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

Public Function CheckListbox(lstBox As ListBox, ByVal strKey As String) As Boolean

    Dim i As Long

    CheckListbox = False

    For i = 0 To lstBox.ListCount - 1

        Dim strCode() As String

        strCode = Split(lstBox.List(i), vbTab, -1, vbBinaryCompare)

        If strCode(0) = strKey Then
            CheckListbox = True
            Exit For

        End If

    Next

End Function

Public Function getHotKeyText() As String

    Dim strKey As String

    strKey = vbNullString

    If FrmSetting.chkShift.value = Checked Then
        strKey = strKey + "Shift + "

    End If

    If FrmSetting.chkCtrl.value = Checked Then
        strKey = strKey + "Ctrl + "

    End If

    If FrmSetting.chkWin.value = Checked Then
        strKey = strKey + "Win + "

    End If

    If FrmSetting.chkAlt.value = Checked Then
        strKey = strKey + "Alt + "

    End If

    getHotKeyText = strKey + FrmSetting.comHKey.Text

End Function

Public Sub LoadSetting()
    Call LoadDatabase '加载数据库

    Dim strSettingFile As String

    strSettingFile = App.Path + "\Config.ini"
    LoadPart0 strSettingFile
    LoadPart1 strSettingFile
    LoadPart3 strSettingFile
    LoadPart4 strSettingFile
    LoadPart6 strSettingFile

End Sub

Public Sub SaveSetting()

    Dim strSettingFile As String

    strSettingFile = App.Path + "\Config.ini"
    SavePart0 strSettingFile
    SavePart1 strSettingFile
    SavePart3 strSettingFile
    SavePart4 strSettingFile
    SavePart6 strSettingFile

End Sub

Private Sub DeleteComboSameItem(cbBox As ComboBox)

    Dim Index As Long, maxCount As Long

    maxCount = cbBox.ListCount - 1

    For Index = 0 To maxCount

        Dim scan As Long

        For scan = Index + 1 To maxCount

            If cbBox.List(Index) = cbBox.List(scan) Then
                cbBox.RemoveItem scan
                scan = scan - 1
                maxCount = maxCount - 1

            End If

        Next
    Next

End Sub

Private Sub DeleteListSameItem(lstBox As ListBox)

    Dim Index As Long, maxCount As Long

    maxCount = lstBox.ListCount - 1

    For Index = 0 To maxCount

        Dim scan As Long

        For scan = Index + 1 To maxCount

            If lstBox.List(Index) = lstBox.List(scan) Then
                lstBox.RemoveItem scan
                scan = scan - 1
                maxCount = maxCount - 1

            End If

        Next
    Next

End Sub

Private Sub LoadDiyCode(lstBox As ListBox, ByVal strSettingFile As String)

    Dim strUserCode As String

    strUserCode = ReadIni("Part6", "DiyCode", strSettingFile)

    If Len(Trim$(strUserCode)) > 0 Then
        lstBox.Clear

        Dim strCode() As String

        strCode = Split(strUserCode, Chr$(1), -1, vbBinaryCompare)

        Dim Index As Long

        For Index = LBound(strCode) To UBound(strCode)

            If Len(Trim$(strCode(Index))) > 0 Then
                lstBox.AddItem Replace$(strCode(Index), Chr$(2), vbTab, 1, -1, vbBinaryCompare)

            End If

        Next
        DeleteListSameItem lstBox
        lstBox.ListIndex = lstBox.ListCount - 1
    Else
        lstBox.Clear

    End If

End Sub

Private Sub LoadPart0(ByVal strSettingFile As String)

    Dim isShift As Boolean, isCtrl As Boolean, isWin As Boolean, isAlt As Boolean

    If ReadIni("Part0", "AutoMini", strSettingFile) = "1" Then
        isAutoMini = True
        FrmSetting.chkAutoMini.value = Checked
    Else
        isAutoMini = False
        FrmSetting.chkAutoMini.value = Unchecked

    End If

    If ReadIni("Part0", "ShowHelp", strSettingFile) = "0" Then
        isShowHelp = False
        FrmSetting.chkShowHelp.value = Unchecked
    Else
        isShowHelp = True
        FrmSetting.chkShowHelp.value = Checked

    End If

    If ReadIni("Part0", "AlwaysOnTop", strSettingFile) = "0" Then
        isAlwaysOnTop = False
        FrmSetting.chkAlwaysOnTop.value = Unchecked
    Else
        isAlwaysOnTop = True
        FrmSetting.chkAlwaysOnTop.value = Checked

    End If

    If ReadIni("Part0", "Shift", strSettingFile) = "1" Then
        isShift = True
        FrmSetting.chkShift.value = Checked
    Else
        isShift = False
        FrmSetting.chkShift.value = Unchecked

    End If

    If ReadIni("Part0", "Ctrl", strSettingFile) = "1" Then
        isCtrl = True
        FrmSetting.chkCtrl.value = Checked
    Else
        isCtrl = False
        FrmSetting.chkCtrl.value = Unchecked

    End If

    If ReadIni("Part0", "Win", strSettingFile) = "0" Then
        isWin = False
        FrmSetting.chkWin.value = Unchecked
    Else
        isWin = True
        FrmSetting.chkWin.value = Checked

    End If

    If ReadIni("Part0", "Alt", strSettingFile) = "1" Then
        isAlt = True
        FrmSetting.chkAlt.value = Checked
    Else
        isAlt = False
        FrmSetting.chkAlt.value = Unchecked

    End If

    HotKeyValue = calcHotKeyValue(isShift, isCtrl, isWin, isAlt)

    Dim strTemp As String

    strTemp = ReadIni("Part0", "Key", strSettingFile)

    If Len(Trim$(strTemp)) > 0 Then

        Dim tmpKey As Integer

        tmpKey = CInt(strTemp)

        If tmpKey <> vbKeyS Then
            If Asc("0") <= tmpKey And tmpKey <= Asc("9") Then
                KeyValue = tmpKey
                FrmSetting.comHKey.ListIndex = KeyValue - Asc("0")
            ElseIf Asc("A") <= tmpKey And tmpKey <= Asc("Z") Then
                KeyValue = tmpKey
                FrmSetting.comHKey.ListIndex = KeyValue - Asc("A") + 10
            Else
                KeyValue = vbKeyS
                FrmSetting.comHKey.ListIndex = KeyValue - Asc("A") + 10

            End If

        Else
            KeyValue = vbKeyS
            FrmSetting.comHKey.ListIndex = KeyValue - Asc("A") + 10

        End If

    Else
        KeyValue = vbKeyS
        FrmSetting.comHKey.ListIndex = KeyValue - Asc("A") + 10

    End If

End Sub

Private Sub LoadPart1(ByVal strSettingFile As String)

    Dim tmpValue As String

    tmpValue = ReadIni("Part1", "TransparentValue", strSettingFile)

    If Len(tmpValue) > 0 Then

        Dim intValue As Integer

        intValue = CInt(tmpValue)

        If intValue <> 192 Then
            If 0 <= intValue And intValue <= 255 Then
                FrmSetting.HScrollTransparent.value = intValue
            Else
                FrmSetting.HScrollTransparent.value = 192

            End If

        Else
            FrmSetting.HScrollTransparent.value = 192

        End If

    Else
        FrmSetting.HScrollTransparent.value = 192

    End If

    If ReadIni("Part1", "Transparent", strSettingFile) = "1" Then
        SetFrmTransparent FrmMain.hwnd, FrmSetting.HScrollTransparent.value
        FrmSetting.chkTransparent.value = Checked
        FrmSetting.HScrollTransparent.Enabled = True
    Else
        UnSetFrmTransparent FrmMain.hwnd
        FrmSetting.chkTransparent.value = Unchecked
        FrmSetting.HScrollTransparent.Enabled = False

    End If

    If ReadIni("Part1", "ShowCode", strSettingFile) = "0" Then
        isShowCode = False
        'FrmMain.lblCode16.Visible = False
        FrmSetting.chkShowCode.value = Unchecked
    Else
        isShowCode = True
        'FrmMain.lblCode16.Visible = True
        FrmSetting.chkShowCode.value = Checked

    End If

    If ReadIni("Part1", "Switch", strSettingFile) = "1" Then
        FrmMain.lblUserCode.Visible = True
        FrmMain.chkAddUserCode.Visible = True
        FrmSetting.chkShowSwitch.value = Checked
    Else
        FrmMain.lblUserCode.Visible = False
        FrmMain.chkAddUserCode.Visible = False
        FrmSetting.chkShowSwitch.value = Unchecked

    End If

    If ReadIni("Part1", "ShowWord", strSettingFile) = "1" Then
        isShowPassword = True
        FrmMain.txtPassword.PasswordChar = vbNullString
        FrmSetting.chkShowPassword.value = Checked
    Else
        isShowPassword = False
        FrmMain.txtPassword.PasswordChar = "*"
        FrmSetting.chkShowPassword.value = Unchecked

    End If

    If ReadIni("Part1", "CloseToHide", strSettingFile) = "0" Then
        isCloseToHide = False
        FrmSetting.optCloseToExit.value = True
    Else
        isCloseToHide = True
        FrmSetting.optClodeToHide.value = True

    End If

End Sub

Private Sub LoadPart3(ByVal strSettingFile As String)

    If ReadIni("Part3", "AutoAddCode", strSettingFile) = "1" Then
        isAutoAddUserCode = True
        FrmSetting.chkAutoAddUserCode.value = Checked
        FrmMain.chkAddUserCode.value = Checked
    Else
        isAutoAddUserCode = False
        FrmSetting.chkAutoAddUserCode.value = Unchecked
        FrmMain.chkAddUserCode.value = Unchecked

    End If

    If ReadIni("Part3", "Prefix", strSettingFile) = "1" Then
        isPrefix = True
        FrmSetting.chkPrefix.value = Checked
    Else
        isPrefix = False
        FrmSetting.chkPrefix.value = Unchecked

    End If

    If ReadIni("Part3", "Suffix", strSettingFile) = "1" Then
        isSuffix = True
        FrmSetting.chkSuffix.value = Checked
    Else
        isSuffix = False
        FrmSetting.chkSuffix.value = Unchecked

    End If

    isUserCodeLoading = True
    LoadUserCode FrmSetting.comPrefixCode, "Prefix", strSettingFile
    LoadUserCode FrmSetting.comSuffixCode, "Suffix", strSettingFile
    isUserCodeLoading = False

End Sub

Private Sub LoadPart4(ByVal strSettingFile As String)

    If ReadIni("Part4", "AutoCopy", strSettingFile) = "0" Then
        isAutoCopy = False
        FrmSetting.chkAutoCopy.value = Unchecked
    Else
        isAutoCopy = True
        FrmSetting.chkAutoCopy.value = Checked

    End If

    If ReadIni("Part4", "AutoUseDomain", strSettingFile) = "0" Then
        isAutoUseDomain = False
        FrmSetting.chkAutoUseDomain.value = Unchecked
    Else
        isAutoUseDomain = True
        FrmSetting.chkAutoUseDomain.value = Checked

    End If

    If ReadIni("Part4", "AutoCheckClipboard", strSettingFile) = "0" Then
        isAutoCheck = False
        FrmSetting.chkAutoCheckClipboard.value = Unchecked
    Else
        isAutoCheck = True
        FrmSetting.chkAutoCheckClipboard.value = Checked

    End If

    If ReadIni("Part4", "DomainSuffix", strSettingFile) = "1" Then
        isDomainSuffix = True
        FrmSetting.chkDomainSuffix.value = Checked
    Else
        isDomainSuffix = False
        FrmSetting.chkDomainSuffix.value = Unchecked

    End If

    If ReadIni("Part4", "Protect", strSettingFile) = "0" Then
        isProtect = False
        FrmSetting.chkProtection.value = Unchecked
    Else
        isProtect = True
        FrmSetting.chkProtection.value = Checked

    End If

    If ReadIni("Part4", "MouseHook", strSettingFile) = "1" Then
        isUseMouseHook = True
        FrmSetting.chkUseMouseHook.value = Checked
    Else
        isUseMouseHook = False
        FrmSetting.chkUseMouseHook.value = Unchecked

    End If

End Sub

Private Sub LoadPart6(ByVal strSettingFile As String)

    If ReadIni("Part6", "DiyWordLength", strSettingFile) = "1" Then
        isDiyWordLength = True
        FrmSetting.chkPasswordLength.value = Checked
        FrmSetting.comPwdLength.Enabled = True
    Else
        isDiyWordLength = False
        FrmSetting.chkPasswordLength.value = Unchecked
        FrmSetting.comPwdLength.Enabled = False

    End If

    Dim strTemp As String

    strTemp = ReadIni("Part6", "WordLength", strSettingFile)

    If Len(Trim$(strTemp)) > 0 Then

        Dim tmpKey As Integer

        tmpKey = CInt(strTemp)

        If tmpKey <> 16 Then
            If 6 <= tmpKey And tmpKey <= 32 Then
                PasswordLength = tmpKey
                FrmSetting.comPwdLength.ListIndex = PasswordLength - 6
            Else
                PasswordLength = 16
                FrmSetting.comPwdLength.ListIndex = 16 - 6

            End If

        Else
            PasswordLength = 16
            FrmSetting.comPwdLength.ListIndex = 16 - 6

        End If

    Else
        PasswordLength = 16
        FrmSetting.comPwdLength.ListIndex = 16 - 6

    End If

    LoadDiyCode FrmSetting.lstDiyKey, strSettingFile

End Sub

Private Sub LoadUserCode(cbBox As ComboBox, _
                         ByVal strFlag As String, _
                         ByVal strSettingFile As String)

    Dim strUserCode As String

    strUserCode = ReadIni("Part3", strFlag + "Code", strSettingFile)

    If Len(Trim$(strUserCode)) > 0 Then
        cbBox.Clear

        Dim strCode() As String

        strCode = Split(strUserCode, Chr$(1), -1, vbBinaryCompare)

        Dim Index As Long

        For Index = LBound(strCode) To UBound(strCode)

            If Len(strCode(Index)) > 0 Then
                cbBox.AddItem strCode(Index)

            End If

        Next
        DeleteComboSameItem cbBox
        strUserCode = ReadIni("Part3", "Last" + strFlag + "Code", strSettingFile)

        If Len(Trim$(strUserCode)) > 0 Then
            Index = CLng(strUserCode)

            If 0 <= Index And Index <= cbBox.ListCount - 1 Then
                cbBox.Text = cbBox.List(Index)

            End If

        End If

    Else
        cbBox.Text = ""

    End If

End Sub

Private Sub SaveDiyCode(lstBox As ListBox, ByVal strSettingFile As String)

    Dim Index As Long, strCode As String, lstCount As Long

    strCode = vbNullString
    lstCount = lstBox.ListCount - 1

    For Index = 0 To lstCount

        Dim strTemp As String

        strTemp = Replace$(lstBox.List(Index), vbTab, Chr$(2), 1, -1, vbBinaryCompare)

        If Index <> lstCount Then
            strCode = strCode + strTemp + Chr$(1)
        Else
            strCode = strCode + strTemp

        End If

    Next
    WriteIni "Part6", "DiyCode", strCode, strSettingFile

End Sub

Private Sub SavePart0(ByVal strSettingFile As String)

    Dim isShift As Boolean, isCtrl As Boolean, isWin As Boolean, isAlt As Boolean

    If FrmSetting.chkAutoMini.value = Checked Then
        isAutoMini = True
        WriteIni "Part0", "AutoMini", "1", strSettingFile
    Else
        isAutoMini = False
        WriteIni "Part0", "AutoMini", "0", strSettingFile

    End If

    If FrmSetting.chkShowHelp.value = Checked Then
        isShowHelp = True
        WriteIni "Part0", "ShowHelp", "1", strSettingFile
    Else
        isShowHelp = False
        WriteIni "Part0", "ShowHelp", "0", strSettingFile

    End If

    If FrmSetting.chkAlwaysOnTop.value = Checked Then
        isAlwaysOnTop = True
        WriteIni "Part0", "AlwaysOnTop", "1", strSettingFile
    Else
        isAlwaysOnTop = False
        WriteIni "Part0", "AlwaysOnTop", "0", strSettingFile

    End If

    If FrmSetting.chkShift.value = Checked Then
        isShift = True
        WriteIni "Part0", "Shift", "1", strSettingFile
    Else
        isShift = False
        WriteIni "Part0", "Shift", "0", strSettingFile

    End If

    If FrmSetting.chkCtrl.value = Checked Then
        isCtrl = True
        WriteIni "Part0", "Ctrl", "1", strSettingFile
    Else
        isCtrl = False
        WriteIni "Part0", "Ctrl", "0", strSettingFile

    End If

    If FrmSetting.chkWin.value = Checked Then
        isWin = True
        WriteIni "Part0", "Win", "1", strSettingFile
    Else
        isWin = False
        WriteIni "Part0", "Win", "0", strSettingFile

    End If

    If FrmSetting.chkAlt.value = Checked Then
        isAlt = True
        WriteIni "Part0", "Alt", "1", strSettingFile
    Else
        isAlt = False
        WriteIni "Part0", "Alt", "0", strSettingFile

    End If

    HotKeyValue = calcHotKeyValue(isShift, isCtrl, isWin, isAlt)
    WriteIni "Part0", "HotKey", HotKeyValue, strSettingFile

    KeyValue = Asc(FrmSetting.comHKey.Text)
    WriteIni "Part0", "Key", KeyValue, strSettingFile

End Sub

Private Sub SavePart1(ByVal strSettingFile As String)

    If FrmSetting.chkTransparent.value = Checked Then
        SetFrmTransparent FrmMain.hwnd, FrmSetting.HScrollTransparent.value
        WriteIni "Part1", "Transparent", "1", strSettingFile
    Else
        UnSetFrmTransparent FrmMain.hwnd
        WriteIni "Part1", "Transparent", "0", strSettingFile

    End If

    WriteIni "Part1", "TransparentValue", FrmSetting.HScrollTransparent.value, strSettingFile

    If FrmSetting.chkShowCode.value = Checked Then
        isShowCode = True
        'FrmMain.lblCode16.Visible = True
        WriteIni "Part1", "ShowCode", "1", strSettingFile
    Else
        isShowCode = False
        'FrmMain.lblCode16.Visible = False
        WriteIni "Part1", "ShowCode", "0", strSettingFile

    End If

    If FrmSetting.chkShowSwitch.value = Checked Then
        FrmMain.lblUserCode.Visible = True
        FrmMain.chkAddUserCode.Visible = True
        WriteIni "Part1", "Switch", "1", strSettingFile
    Else
        FrmMain.lblUserCode.Visible = False
        FrmMain.chkAddUserCode.Visible = False
        WriteIni "Part1", "Switch", "0", strSettingFile

    End If

    If FrmSetting.chkShowPassword.value = Checked Then
        isShowPassword = True
        FrmMain.txtPassword.PasswordChar = vbNullString
        WriteIni "Part1", "ShowWord", "1", strSettingFile
    Else
        isShowPassword = False
        FrmMain.txtPassword.PasswordChar = "*"
        WriteIni "Part1", "ShowWord", "0", strSettingFile

    End If

    If FrmSetting.optClodeToHide.value = True Then
        isCloseToHide = True
        WriteIni "Part1", "CloseToHide", "1", strSettingFile
    Else
        isCloseToHide = False
        WriteIni "Part1", "CloseToHide", "0", strSettingFile

    End If

End Sub

Private Sub SavePart3(ByVal strSettingFile As String)

    If FrmSetting.chkAutoAddUserCode.value = Checked Then
        FrmMain.chkAddUserCode.value = Checked
        isAutoAddUserCode = True
        WriteIni "Part3", "AutoAddCode", "1", strSettingFile
    Else
        FrmMain.chkAddUserCode.value = Unchecked
        isAutoAddUserCode = False
        WriteIni "Part3", "AutoAddCode", "0", strSettingFile

    End If

    If FrmSetting.chkPrefix.value = Checked Then
        isPrefix = True
        WriteIni "Part3", "Prefix", "1", strSettingFile
    Else
        isPrefix = False
        WriteIni "Part3", "Prefix", "0", strSettingFile

    End If

    If FrmSetting.chkSuffix.value = Checked Then
        isSuffix = True
        WriteIni "Part3", "Suffix", "1", strSettingFile
    Else
        isSuffix = False
        WriteIni "Part3", "Suffix", "0", strSettingFile

    End If

    AddStrToCombox FrmSetting.comPrefixCode, FrmSetting.comPrefixCode.Text
    SaveUserCode FrmSetting.comPrefixCode, "Prefix", strSettingFile
    AddStrToCombox FrmSetting.comSuffixCode, FrmSetting.comSuffixCode.Text
    SaveUserCode FrmSetting.comSuffixCode, "Suffix", strSettingFile

End Sub

Private Sub SavePart4(ByVal strSettingFile As String)

    If FrmSetting.chkAutoCopy.value = Checked Then
        isAutoCopy = True
        WriteIni "Part4", "AutoCopy", "1", strSettingFile
    Else
        isAutoCopy = False
        WriteIni "Part4", "AutoCopy", "0", strSettingFile

    End If

    If FrmSetting.chkAutoUseDomain.value = Checked Then
        isAutoUseDomain = True
        WriteIni "Part4", "AutoUseDomain", "1", strSettingFile
    Else
        isAutoUseDomain = False
        WriteIni "Part4", "AutoUseDomain", "0", strSettingFile

    End If

    If FrmSetting.chkAutoCheckClipboard.value = Checked Then
        isAutoCheck = True
        WriteIni "Part4", "AutoCheckClipboard", "1", strSettingFile
    Else
        isAutoCheck = False
        WriteIni "Part4", "AutoCheckClipboard", "0", strSettingFile

    End If

    If FrmSetting.chkDomainSuffix.value = Checked Then
        isDomainSuffix = True
        WriteIni "Part4", "DomainSuffix", "1", strSettingFile
    Else
        isDomainSuffix = False
        WriteIni "Part4", "DomainSuffix", "0", strSettingFile

    End If

    If FrmSetting.chkProtection.value = Checked Then
        isProtect = True
        WriteIni "Part4", "Protect", "1", strSettingFile
    Else
        isProtect = False
        WriteIni "Part4", "Protect", "0", strSettingFile

    End If

    If FrmSetting.chkUseMouseHook.value = Checked Then
        isUseMouseHook = True
        WriteIni "Part4", "MouseHook", "1", strSettingFile
    Else
        isUseMouseHook = False
        WriteIni "Part4", "MouseHook", "0", strSettingFile

    End If

End Sub

Private Sub SavePart6(ByVal strSettingFile As String)

    If FrmSetting.chkPasswordLength.value = Checked Then
        isDiyWordLength = True
        PasswordLength = FrmSetting.comPwdLength.Text
        WriteIni "Part6", "DiyWordLength", "1", strSettingFile
    Else
        isDiyWordLength = False
        PasswordLength = 16
        WriteIni "Part6", "DiyWordLength", "0", strSettingFile

    End If

    WriteIni "Part6", "WordLength", PasswordLength, strSettingFile
    SaveDiyCode FrmSetting.lstDiyKey, strSettingFile

End Sub

Private Sub SaveUserCode(cbBox As ComboBox, _
                         ByVal strFlag As String, _
                         ByVal strSettingFile As String)

    Dim Index As Long, strCode As String, lstCount As Long

    strCode = vbNullString
    lstCount = cbBox.ListCount - 1

    For Index = 0 To lstCount

        If Index <> lstCount Then
            strCode = strCode + cbBox.List(Index) + Chr$(1)
        Else
            strCode = strCode + cbBox.List(Index)

        End If

    Next
    WriteIni "Part3", strFlag + "Code", strCode, strSettingFile

    If cbBox.ListIndex < 0 Then
        WriteIni "Part3", "Last" + strFlag + "Code", cbBox.ListCount - 1, strSettingFile
    Else
        WriteIni "Part3", "Last" + strFlag + "Code", cbBox.ListIndex, strSettingFile

    End If

End Sub
