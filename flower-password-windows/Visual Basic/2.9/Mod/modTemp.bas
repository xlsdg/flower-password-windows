Attribute VB_Name = "modTemp"
Option Explicit

Private Declare Function GetTempPath _
                Lib "kernel32.dll" _
                Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                      ByVal lpBuffer As String) As Long

Public Function GetTempPathByApi() As String

    Dim tempPath As String, lenPath As Long

    tempPath = String$(255, 0)
    lenPath = GetTempPath(256, tempPath)
    GetTempPathByApi = Left$(tempPath, lenPath)

End Function
