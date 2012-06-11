Attribute VB_Name = "modXPStyle"
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
