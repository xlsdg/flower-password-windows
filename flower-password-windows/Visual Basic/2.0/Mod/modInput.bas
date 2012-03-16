Attribute VB_Name = "modInput"
Option Explicit

Public Password_Hwnd As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function MapVirtualKey _
                Lib "user32" _
                Alias "MapVirtualKeyA" (ByVal wCode As Long, _
                                        ByVal wMapType As Long) As Long

Private Declare Sub keybd_event _
                Lib "user32" (ByVal bVk As Byte, _
                              ByVal bScan As Byte, _
                              ByVal dwFlags As Long, _
                              ByVal dwExtraInfo As Long)

Private Const KEYEVENTF_KEYDOWN = &H0

Private Const KEYEVENTF_KEYUP = &H2

Private Declare Function SendInput _
                Lib "user32.dll" (ByVal nInputs As Long, _
                                  pInputs As GENERALINPUT, _
                                  ByVal cbSize As Long) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
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

Private Sub SendInputLowerKey(ByVal bkey As Long)

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
    SendInput 2, GInput(0), Len(GInput(0))

End Sub

Private Sub SendInputUpperKey(ByVal bkey As Long)

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
    SendInput 4, GInput(0), Len(GInput(0))

End Sub

Public Sub PostCode(ByVal TextBoxHwnd As Long, ByVal strCode As String)

    Dim code_len As Long

    code_len = Len(strCode)

    If TextBoxHwnd > 0 And code_len > 0 Then

        'SetFocus TextBoxHwnd
        SetCapsLock False
        Dim i As Long

        For i = 1 To code_len

            Dim key_code As Integer

            key_code = Asc(Mid$(strCode, i, 1))

            If Asc("0") <= key_code And key_code <= Asc("9") Then
                InputNumber key_code
            ElseIf Asc("A") <= key_code And key_code <= Asc("Z") Then
                InputUpperCase key_code
            ElseIf Asc("a") <= key_code And key_code <= Asc("z") Then
                InputLowerCase (key_code)

            End If

        Next

    End If

End Sub

Private Sub InputLowerCase(ByVal keycode As Integer)
    keycode = keycode - Asc("a") + Asc("A")
    'keybd_event keycode, MapVirtualKey(keycode, 0), KEYEVENTF_KEYDOWN, 0
    'keybd_event keycode, MapVirtualKey(keycode, 0), KEYEVENTF_KEYUP, 0
    SendInputLowerKey keycode

End Sub

Private Sub InputUpperCase(ByVal keycode As Integer)
    'keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), KEYEVENTF_KEYDOWN, 0
    'keybd_event keycode, MapVirtualKey(keycode, 0), KEYEVENTF_KEYDOWN, 0
    'keybd_event keycode, MapVirtualKey(keycode, 0), KEYEVENTF_KEYUP, 0
    'keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), KEYEVENTF_KEYUP, 0
    SendInputUpperKey keycode

End Sub

Private Sub InputNumber(ByVal keycode As Integer)
    'keybd_event keycode, MapVirtualKey(keycode, 0), KEYEVENTF_KEYDOWN, 0
    'keybd_event keycode, MapVirtualKey(keycode, 0), KEYEVENTF_KEYUP, 0
    SendInputLowerKey keycode

End Sub
