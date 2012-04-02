Attribute VB_Name = "modXPStyle"
Option Explicit

Private Declare Function InitCommonControlsEx _
                Lib "Comctl32.dll" (Iccex As TagInitCommonControlsEx) As Boolean

Private Type TagInitCommonControlsEx

    LngSize As Long
    LngICC As Long

End Type

Private Const ICC_USEREX_CLASSES As Long = &H200

Public Function InitCommonControlsVB() As Boolean

    On Error Resume Next

    Dim Iccex As TagInitCommonControlsEx

    With Iccex
        .LngSize = LenB(Iccex)
        .LngICC = ICC_USEREX_CLASSES

    End With

    InitCommonControlsEx Iccex
    InitCommonControlsVB = (Err.Number = 0)

End Function
