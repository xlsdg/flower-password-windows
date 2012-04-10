Attribute VB_Name = "modMemory"
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

Private Declare Function SetProcessWorkingSetSize _
                Lib "kernel32.dll" (ByVal hProcess As Long, _
                                    ByVal dwMinimumWorkingSetSize As Long, _
                                    ByVal dwMaximumWorkingSetSize As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long

Public Function ZipMemory() As Long
    ZipMemory = SetProcessWorkingSetSize(GetCurrentProcess, -1, -1)

End Function
