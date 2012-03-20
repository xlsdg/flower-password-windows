Attribute VB_Name = "modIEObject"
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint _
                Lib "user32" (ByVal xPoint As Long, _
                              ByVal yPoint As Long) As Long

Private Declare Function ScreenToClient _
                Lib "user32" (ByVal hwnd As Long, _
                              lpPoint As POINTAPI) As Long

Private Type POINTAPI

    x As Long
    y As Long

End Type

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
                Lib "user32" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Private Declare Function EnumChildWindows _
                Lib "user32" (ByVal hWndParent As Long, _
                              ByVal lpEnumFunc As Long, _
                              lParam As Long) As Long

Private Declare Function RegisterWindowMessage _
                Lib "user32" _
                Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Function SendMessageTimeout _
                Lib "user32" _
                Alias "SendMessageTimeoutA" (ByVal hwnd As Long, _
                                             ByVal msg As Long, _
                                             ByVal wParam As Long, _
                                             lParam As Any, _
                                             ByVal fuFlags As Long, _
                                             ByVal uTimeout As Long, _
                                             lpdwResult As Long) As Long

Private Const SMTO_ABORTIFHUNG = &H2

Private Declare Function ObjectFromLresult _
                Lib "oleacc" (ByVal lResult As Long, _
                              riid As UUID, _
                              ByVal wParam As Long, _
                              ppvObject As Any) As Long

Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long
  
'
'   IEDOMFromhWnd
'
'   Returns   the   IHTMLDocument   interface   from   a   WebBrowser   window
'
'   hWnd   -   Window   handle   of   the   control
'
Private Function IEDOMFromhWnd(ByVal hwnd As Long) As IHTMLDocument

    Dim IID_IHTMLDocument As UUID

    Dim hWndChild         As Long

    Dim lRes              As Long

    Dim lMsg              As Long

    Dim hr                As Long

    If hwnd <> 0 Then
        If Not IsIEServerWindow(hwnd) Then
            '   Find   a   child   IE   server   window
            EnumChildWindows hwnd, AddressOf EnumChildProc, hwnd

        End If

        If hwnd <> 0 Then
            '   Register   the   message
            lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")
            '   Get   the   object   pointer
            Call SendMessageTimeout(hwnd, lMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes)

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
                hr = ObjectFromLresult(lRes, IID_IHTMLDocument, 0, IEDOMFromhWnd)

            End If

        End If

    End If

End Function
  
Private Function IsIEServerWindow(ByVal hwnd As Long) As Boolean

    Dim lRes       As Long

    Dim sClassName As String

    '   Initialize   the   buffer
    sClassName = String$(100, 0)
    '   Get   the   window   class   name
    lRes = GetClassName(hwnd, sClassName, Len(sClassName))
    sClassName = Left$(sClassName, lRes)
    IsIEServerWindow = InStr(1, sClassName, "Internet Explorer_Server", vbTextCompare)

End Function
  
'
'   Copy   this   function   to   a   .bas   module
'
Private Function EnumChildProc(ByVal hwnd As Long, lParam As Long) As Long

    If IsIEServerWindow(hwnd) Then
        lParam = hwnd
    Else
        EnumChildProc = 1

    End If

End Function

Public Function PostCodeToIE(ByVal strCode As String) As Long

    Dim pt As POINTAPI, IE_hWnd As Long

    GetCursorPos pt
    IE_hWnd = WindowFromPoint(pt.x, pt.y)

    If IE_hWnd > 0 Then
        ScreenToClient IE_hWnd, pt

        If SetCodeInIE(IE_hWnd, strCode, pt.x, pt.y) > 0 Then
            PostCodeToIE = 1
        Else
            PostCodeToIE = 0

        End If

    Else
        PostCodeToIE = 0

    End If

End Function

Private Function SetCodeInIE(ByVal hwnd As Long, _
                             ByVal strCode As String, _
                             ByVal cx As Long, _
                             ByVal cy As Long) As Long

    Dim Doc As IHTMLDocument

    Set Doc = IEDOMFromhWnd(hwnd)

    If Doc Is Nothing Then
        SetCodeInIE = 0
    Else

        Dim Ele As IHTMLElement

        Set Ele = Doc.elementFromPoint(cx, cy)

        If Ele Is Nothing Then
            SetCodeInIE = 0
        Else

            If Ele.Type = "password" Then
                Ele.value = strCode
                SetCodeInIE = 1
            Else
                SetCodeInIE = 0

            End If

        End If

    End If

End Function
