Attribute VB_Name = "modIEObject"
Option Explicit

'
'   Requires:   reference   to   "Microsoft   HTML   Object   Library"
'
Private Type UUID

    Data1   As Long
    Data2   As Integer
    Data3   As Integer
    Data4(0 To 7)       As Byte

End Type

Private Declare Function GetClassName _
                Lib "user32.dll" _
                Alias "GetClassNameA" (ByVal hWnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Private Declare Function RegisterWindowMessage _
                Lib "user32.dll" _
                Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Function SendMessageTimeout _
                Lib "user32.dll" _
                Alias "SendMessageTimeoutA" (ByVal hWnd As Long, _
                                             ByVal msg As Long, _
                                             ByVal wParam As Long, _
                                             lParam As Any, _
                                             ByVal fuFlags As Long, _
                                             ByVal uTimeout As Long, _
                                             lpdwResult As Long) As Long

Private Const SMTO_ABORTIFHUNG = &H2

Private Declare Function ObjectFromLresult _
                Lib "oleacc.dll" (ByVal lResult As Long, _
                                  riid As UUID, _
                                  ByVal wParam As Long, _
                                  ppvObject As Any) As Long

Public Function GetIEDomainName(ByVal hWnd As Long) As String

    Dim Doc As IHTMLDocument

    Set Doc = IEDOMFromhWnd(hWnd)

    If Doc Is Nothing Then
        GetIEDomainName = vbNullString
    Else
        GetIEDomainName = Doc.domain

    End If

    Set Doc = Nothing

End Function

Public Function isInternetExplorer(ByVal WinWnd As Long) As Boolean

    If WinWnd > 0 Then

        Dim RetVal As Long, lpClassName As String

        lpClassName = Space$(256)
        RetVal = GetClassName(WinWnd, lpClassName, 256)

        If InStr(Left$(lpClassName, RetVal), "Internet Explorer_Server") > 0 Then
            isInternetExplorer = True
        Else
            isInternetExplorer = False

        End If

    Else
        isInternetExplorer = False

    End If

End Function

Public Function PostCodeToIE(ByVal hWnd As Long, ByVal strCode As String) As Long

    Dim Doc As IHTMLDocument

    Set Doc = IEDOMFromhWnd(hWnd)

    If Doc Is Nothing Then
        PostCodeToIE = 0
    Else
        PostCodeToIE = 0

        Dim objInput As Object

        For Each objInput In Doc.getElementsByTagName("INPUT")

            If StrComp(objInput.Type, "password", 1) = 0 Then
                objInput.value = strCode
                PostCodeToIE = 1
                Exit For

            End If

        Next
        Set objInput = Nothing

    End If

    Set Doc = Nothing

End Function

'
'   IEDOMFromhWnd
'
'   Returns   the   IHTMLDocument   interface   from   a   WebBrowser   window
'
'   hWnd   -   Window   handle   of   the   control
'
Private Function IEDOMFromhWnd(ByVal hWnd As Long) As IHTMLDocument

    If hWnd <> 0 Then

        Dim IID_IHTMLDocument As UUID

        Dim lRes              As Long

        Dim lMsg              As Long

        '   Register   the   message
        lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")
        '   Get   the   object   pointer
        Call SendMessageTimeout(hWnd, lMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes)

        If lRes Then

            '   Initialize   the   interface   ID
            With IID_IHTMLDocument
                .Data1 = &H626FC520
                .Data2 = &HA41E
                .Data3 = &H11CF
                .Data4(0) = &HA7
                .Data4(1) = &H31
                .Data4(2) = &H0
                .Data4(3) = &HA0
                .Data4(4) = &HC9
                .Data4(5) = &H8
                .Data4(6) = &H26
                .Data4(7) = &H37

            End With

            '   Get   the   object   from   lRes
            Call ObjectFromLresult(lRes, IID_IHTMLDocument, 0, IEDOMFromhWnd)

        End If

    End If

End Function
