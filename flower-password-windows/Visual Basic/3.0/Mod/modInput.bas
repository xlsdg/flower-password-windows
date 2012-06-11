Attribute VB_Name = "modInput"
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

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Function PostMessage _
                Lib "user32.dll" _
                Alias "PostMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function MapVirtualKey _
                Lib "user32.dll" _
                Alias "MapVirtualKeyA" (ByVal wCode As Long, _
                                        ByVal wMapType As Long) As Long

Private Declare Function SendMessage _
                Lib "user32.dll" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const WM_SYSKEYDOWN = &H104

Private Const WM_SYSKEYUP = &H105

Private Const WM_KEYDOWN = &H100

Private Const WM_KEYUP = &H101

Private Const WM_CHAR = &H102

Private Const EM_REPLACESEL = &HC2

Private Const WM_SETTEXT = &HC

Private Const WM_COPY = &H301

Private Const WM_CUT = &H300

Private Const WM_PASTE = &H302

Private Const WM_GETTEXT = &HD

Private Const WM_GETTEXTLENGTH = &HE

Private Const KEYEVENTF_KEYDOWN = &H0

Private Const KEYEVENTF_KEYUP = &H2

Private Declare Function SendInput _
                Lib "user32.dll" (ByVal nInputs As Long, _
                                  pInputs As GENERALINPUT, _
                                  ByVal cbSize As Long) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32.dll" _
                Alias "RtlMoveMemory" (pDst As Any, _
                                       pSrc As Any, _
                                       ByVal ByteLen As Long)

Private Type GENERALINPUT

    dwType As Long
    xi(0 To 23) As Byte

End Type

Private Type KEYBDINPUT

    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long

End Type

Private Const INPUT_KEYBOARD = &H1

Public Function PoseCodeToClipboard(ByVal strCode As String) As Long

        '<EhHeader>
        On Error GoTo PoseCodeToClipboard_Err

        '</EhHeader>
100     If isAutoCopy Then
102         Clipboard.Clear
104         Clipboard.SetText strCode, vbCFText
106         PoseCodeToClipboard = 1
        Else
108         PoseCodeToClipboard = 0

        End If

        '<EhFooter>
        Exit Function
PoseCodeToClipboard_Err:
        PoseCodeToClipboard = 0
        'MsgBox Err.Description & vbCrLf & _
         "in FlowerPassword.modInput.PoseCodeToClipboard " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>
End Function

Public Function PostCode(ByVal strCode As String, ByVal TextBoxHwnd As Long) As Long

    If isInternetExplorer(TextBoxHwnd) Then
        If PostCodeToIE(TextBoxHwnd, strCode) = 1 Then
            PostCode = 1
        Else
            Sleep 750
            PoseCodeBySendInput strCode
            PostCode = 2

        End If

    ElseIf isChrome(TextBoxHwnd) Then
        Sleep 750
        PoseCodeBySendInput strCode
        PostCode = 3
    ElseIf isFirefox(TextBoxHwnd) Then
        Sleep 750
        PoseCodeBySendInput strCode
        PostCode = 4
    ElseIf isOpera(TextBoxHwnd) Then
        Sleep 750
        PoseCodeBySendInput strCode
        PostCode = 5
    ElseIf isMaxthon(TextBoxHwnd) Then
        Sleep 750
        PoseCodeBySendInput strCode
        PostCode = 6
    Else

        If TextBoxHwnd > 0 Then
            'PoseCodeByKeyDown TextBoxHwnd, strCode
            PoseCodeByPaste TextBoxHwnd
            'PoseCodeBySetText TextBoxHwnd, strCode
            'If GetCodeLength(TextBoxHwnd) <> 16 Then
            'PoseCodeBySendInput strCode
            'End If
            PostCode = 7
        Else
            Sleep 750
            PoseCodeBySendInput strCode
            PostCode = 8

        End If

    End If

End Function

Private Function GetCodeLength(ByVal TextBoxHwnd As Long) As Long
    GetCodeLength = SendMessage(TextBoxHwnd, WM_GETTEXTLENGTH, 0, vbNull)

End Function

Private Function MakeKeyLparam(ByVal VirtualKey As Long, ByVal flag As Long) As Long

    Dim Firstbyte As String    'lparam参数的24-31位

    If flag = WM_KEYDOWN Then  '如果是按下键
        Firstbyte = "00"
    Else
        Firstbyte = "C0"       '如果是释放键

    End If

    Dim Scancode As Long

    '获得键的扫描码
    Scancode = MapVirtualKey(VirtualKey, 0)

    Dim Secondbyte As String   'lparam参数的16-23位，即虚拟键扫描码

    Secondbyte = Right$("00" & Hex$(Scancode), 2)

    Dim s As String

    s = Firstbyte & Secondbyte & "0001"  '0001为lparam参数的0-15位，即发送次数和其它扩展信息
    MakeKeyLparam = CLng("&H" & s)

End Function

Private Function PoseCodeByKeyDown(ByVal TextBoxHwnd As Long, _
                                   ByVal strCode As String) As Long

    Dim code_len As Long

    code_len = Len(strCode)

    If Len(code_len) > 0 Then

        Dim i As Long, result As Long

        For i = 1 To code_len

            Dim key_code As String

            result = 0
            key_code = Mid$(strCode, i, 1)
            result = PostMessage(TextBoxHwnd, WM_KEYDOWN, Asc(UCase$(key_code)), MakeKeyLparam(Asc(UCase$(key_code)), WM_KEYDOWN))
            result = PostMessage(TextBoxHwnd, WM_CHAR, Asc(key_code), MakeKeyLparam(Asc(UCase$(key_code)), WM_KEYDOWN))
            result = PostMessage(TextBoxHwnd, WM_KEYUP, Asc(UCase$(key_code)), MakeKeyLparam(Asc(UCase$(key_code)), WM_KEYUP))

            If result = 0 Then
                PoseCodeByKeyDown = 0
                Exit For

            End If

        Next
        PoseCodeByKeyDown = result

    End If

End Function

Private Function PoseCodeByPaste(ByVal TextBoxHwnd As Long) As Long
    PoseCodeByPaste = SendMessage(TextBoxHwnd, WM_PASTE, 0, 0)

End Function

Private Function PoseCodeBySendInput(ByVal strCode As String) As Long

    Dim code_len As Long

    code_len = Len(strCode)

    If Len(code_len) > 0 Then
        SetCapsLock False

        Dim i As Long, result As Long

        For i = 1 To code_len

            Dim key_code As Integer

            result = 0
            key_code = Asc(Mid$(strCode, i, 1))

            If Asc("0") <= key_code And key_code <= Asc("9") Then
                result = SendInputNumber(key_code)
            ElseIf Asc("A") <= key_code And key_code <= Asc("Z") Then
                result = SendInputUpperCase(key_code)
            ElseIf Asc("a") <= key_code And key_code <= Asc("z") Then
                result = SendInputLowerCase(key_code)
            Else
                PoseCodeBySendInput = 0
                Exit For

            End If

            If result = 0 Then
                PoseCodeBySendInput = 0
                Exit For

            End If

        Next
        PoseCodeBySendInput = result

    End If

End Function

Private Function PoseCodeBySetText(ByVal TextBoxHwnd As Long, _
                                   ByVal strCode As String) As Long
    PoseCodeBySetText = SendMessage(TextBoxHwnd, WM_SETTEXT, 0, ByVal strCode)

End Function

Private Function SendInputLowerCase(ByVal KeyCode As Integer) As Long
    KeyCode = KeyCode - Asc("a") + Asc("A")
    SendInputLowerCase = SendInputLowerKey(KeyCode)

End Function

Private Function SendInputLowerKey(ByVal bkey As Long) As Long

    Dim GInput(0 To 1) As GENERALINPUT

    Dim KInput         As KEYBDINPUT

    KInput.wVk = bkey
    KInput.dwFlags = KEYEVENTF_KEYDOWN
    GInput(0).dwType = INPUT_KEYBOARD
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    KInput.wVk = bkey
    KInput.dwFlags = KEYEVENTF_KEYUP
    GInput(1).dwType = INPUT_KEYBOARD
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    SendInputLowerKey = SendInput(2, GInput(0), Len(GInput(0)))

End Function

Private Function SendInputNumber(ByVal KeyCode As Integer) As Long
    SendInputNumber = SendInputLowerKey(KeyCode)

End Function

Private Function SendInputUpperCase(ByVal KeyCode As Integer) As Long
    SendInputUpperCase = SendInputUpperKey(KeyCode)

End Function

Private Function SendInputUpperKey(ByVal bkey As Long) As Long

    Dim GInput(0 To 3) As GENERALINPUT

    Dim KInput         As KEYBDINPUT

    KInput.wVk = vbKeyShift
    KInput.dwFlags = KEYEVENTF_KEYDOWN
    GInput(0).dwType = INPUT_KEYBOARD
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    KInput.wVk = bkey
    KInput.dwFlags = KEYEVENTF_KEYDOWN
    GInput(1).dwType = INPUT_KEYBOARD
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    KInput.wVk = bkey
    KInput.dwFlags = KEYEVENTF_KEYUP
    GInput(2).dwType = INPUT_KEYBOARD
    CopyMemory GInput(2).xi(0), KInput, Len(KInput)
    KInput.wVk = vbKeyShift
    KInput.dwFlags = KEYEVENTF_KEYUP
    GInput(3).dwType = INPUT_KEYBOARD
    CopyMemory GInput(3).xi(0), KInput, Len(KInput)
    SendInputUpperKey = SendInput(4, GInput(0), Len(GInput(0)))

End Function

