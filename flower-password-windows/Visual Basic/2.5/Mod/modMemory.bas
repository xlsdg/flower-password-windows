Attribute VB_Name = "modMemory"
Option Explicit

Private Declare Function SetProcessWorkingSetSize _
                Lib "kernel32.dll" (ByVal hProcess As Long, _
                                    ByVal dwMinimumWorkingSetSize As Long, _
                                    ByVal dwMaximumWorkingSetSize As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long

Public Function ZipMemory() As Long
    ZipMemory = SetProcessWorkingSetSize(GetCurrentProcess, -1, -1)

End Function
