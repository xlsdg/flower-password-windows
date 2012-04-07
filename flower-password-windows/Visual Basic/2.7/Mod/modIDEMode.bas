Attribute VB_Name = "modIDEMode"
Option Explicit

Private Declare Function GetModuleFileName _
                Lib "kernel32.dll" _
                Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                            ByVal lpFileName As String, _
                                            ByVal nSize As Long) As Long

Public Function isRunInIDEMode() As Boolean

    Dim strFileName As String, lngCount As Long, strAppName As String, lngAppName As Long

    strFileName = String$(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left$(strFileName, lngCount)
    strAppName = App.EXEName & ".exe"
    lngAppName = Len(strAppName)

    If UCase$(Right$(strFileName, lngAppName)) = UCase$(strAppName) Then
        isRunInIDEMode = False
    Else
        isRunInIDEMode = True

    End If

End Function
