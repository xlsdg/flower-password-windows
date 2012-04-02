Attribute VB_Name = "modIniRW"
Option Explicit

Private Declare Function GetPrivateProfileString _
                Lib "kernel32.dll" _
                Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                  lpKeyName As Any, _
                                                  ByVal lpDefault As String, _
                                                  ByVal lpRetunedString As String, _
                                                  ByVal nSize As Long, _
                                                  ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString _
                Lib "kernel32.dll" _
                Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lplFileName As String) As Long

Private Const MAX_PATH = 260

Public Function ReadIni(ByVal AppName As String, _
                        ByVal KeyName As String, _
                        ByVal FileName As String) As String

    Dim returnBuffer As String, lpRetStr As String

    returnBuffer = Space$(MAX_PATH)  '填充缓冲区
    lpRetStr = GetPrivateProfileString(ByVal AppName, ByVal KeyName, vbNullString, ByVal returnBuffer, ByVal Len(returnBuffer), ByVal FileName) '返回复制到缓冲区里的字节数目
    ReadIni = Left$(returnBuffer, lpRetStr)  '得到字符串

End Function

Public Function WriteIni(ByVal lpApplicationName As String, _
                         ByVal lpKeyName As String, _
                         ByVal lpString As String, _
                         ByVal lplFileName As String) As Long
    WriteIni = WritePrivateProfileString(lpApplicationName, lpKeyName, lpString, lplFileName)

End Function
