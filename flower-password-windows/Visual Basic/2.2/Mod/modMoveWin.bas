Attribute VB_Name = "modMoveWin"
Option Explicit

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Declare Function SendMessage _
                Lib "user32.dll" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const WM_SYSCOMMAND = &H112

Private Const SC_MOVE = &HF010&

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Public Function SetWinMove(ByVal WinhWnd As Long) As Long
    ReleaseCapture
    SetWinMove = SendMessage(WinhWnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0&)

    'SendMessage WinHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function
