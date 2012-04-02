Attribute VB_Name = "modOpenWeb"
Option Explicit

Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Private Declare Function ShellExecute _
                Lib "Shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long
                                       
Public Function OpenWebsite(ByVal strUrl As String) As Long
    OpenWebsite = ShellExecute(GetForegroundWindow, "Open", strUrl, 0, 0, 0)

End Function
