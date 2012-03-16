Attribute VB_Name = "modCapsLock"
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Function MapVirtualKey _
               Lib "user32" _
               Alias "MapVirtualKeyA" (ByVal wCode As Long, _
                                       ByVal wMapType As Long) As Long

Private Declare Sub keybd_event _
               Lib "user32" (ByVal bVk As Byte, _
                             ByVal bScan As Byte, _
                             ByVal dwFlags As Long, _
                             ByVal dwExtraInfo As Long)

Private Const KEYEVENTF_EXTENDEDKEY = &H1

Private Const KEYEVENTF_KEYUP = &H2

Public Sub SetCapsLock(ByVal bLock As Boolean)

    Dim Check As Boolean, ScanCode As Long

    Check = CBool(GetKeyState(vbKeyCapital))

    If Check <> bLock Then
        ScanCode = MapVirtualKey(vbKeyCapital, 0)
        Call keybd_event(vbKeyCapital, ScanCode, 0, 0)
        Call keybd_event(vbKeyCapital, ScanCode, KEYEVENTF_KEYUP, 0)

    End If

End Sub
