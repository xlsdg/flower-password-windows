Attribute VB_Name = "modExplorer"
Option Explicit

Public Function FilterDomainName(ByVal Str_DomainName As String) As String

    Dim Str_Extensions As String

    Str_Extensions = ".com.cn|.net.cn|.gov.cn|.edu.cn|.org.cn|.mil.cn|.com.hk|.travel|.ac.cn|.bj.cn|.sh.cn|.tj.cn|.cq.cn|.he.cn|.sx.cn|.nm.cn|.ln.cn|.jl.cn|.hl.cn|.js.cn|.zj.cn|.ah.cn|.fj.cn|.jx.cn|.sd.cn|.ha.cn|.hb.cn|.hn.cn|.gd.cn|.gx.cn|.hi.cn|.sc.cn|.gz.cn|.yn.cn|.xz.cn|.sn.cn|.gs.cn|.qh.cn|.nx.cn|.xj.cn|.tw.cn|.hk.cn|.mo.cn|.info|.mobi|.name|.asia|.com|.net|.org|.biz|.tel|.mil|.int|.edu|.hk|.me|.us|.cc|.tv"

    Dim Arr_Ext() As String

    Arr_Ext = Split(Str_Extensions, "|")
    Str_DomainName = LCase$(Str_DomainName)

    Dim X As Integer

    For X = LBound(Arr_Ext) To UBound(Arr_Ext)

        Dim arr_len As Long, str_len As Long

        arr_len = Len(Arr_Ext(X))
        str_len = Len(Str_DomainName)

        If Right$(Str_DomainName, arr_len) = Arr_Ext(X) And str_len > arr_len Then
            Str_DomainName = Left$(Str_DomainName, str_len - arr_len)
            str_len = Len(Str_DomainName)

            Dim v As Integer

            v = InStrRev(Str_DomainName, ".")

            If v = 0 Then
                FilterDomainName = Str_DomainName
            Else
                FilterDomainName = Right$(Str_DomainName, str_len - v)

            End If

            If isDomainSuffix Then
                FilterDomainName = FilterDomainName + Arr_Ext(X)

            End If

            Exit For

        End If

    Next

End Function

Public Function GetWebsiteName(ByVal strUrl As String) As String
    strUrl = LCase$(strUrl)

    Dim a As Long

    a = InStr(strUrl, "//")

    If a > 0 Then
        strUrl = Right$(strUrl, Len(strUrl) - a - 1)

    End If

    a = InStr(strUrl, "/")

    If a > 0 Then
        strUrl = Left$(strUrl, a - 1)

    End If

    GetWebsiteName = strUrl

End Function

Public Function isClipboardAsUrl() As String

    If Clipboard.GetFormat(vbCFText) Then

        Dim str_url As String, str_len As Long

        str_url = LCase$(Clipboard.GetText)
        str_len = Len(str_url)

        If str_len > 0 Then
            isClipboardAsUrl = vbNullString

            Dim Str_Sites As String

            Str_Sites = LCase$("http|https|ftp|mms|rtsp|rtmp|mmst|bt|www.|ftp.|pop.|smtp.|wap.|m.|3g.")

            Dim Arr_Ext() As String

            Arr_Ext = Split(Str_Sites, "|")

            Dim X As Integer

            For X = LBound(Arr_Ext) To UBound(Arr_Ext)

                Dim arr_len As Long

                arr_len = Len(Arr_Ext(X))

                If Left$(str_url, arr_len) = Arr_Ext(X) And str_len > arr_len Then
                    isClipboardAsUrl = GetWebsiteName(str_url)
                    Exit For

                End If

            Next
        Else
            isClipboardAsUrl = vbNullString

        End If

    Else
        isClipboardAsUrl = vbNullString

    End If

End Function

Public Function SetUrlAsKey(ByVal hwnd As Long) As Long

    Dim strUrl As String

    If isInternetExplorer(hwnd) Then
        strUrl = GetIEDomainName(hwnd)

        If Len(strUrl) > 0 Then
            FrmMain.comKey.Text = FilterDomainName(strUrl)
            SetUrlAsKey = 1
        Else
            SetUrlAsKey = 0

        End If

    ElseIf isChrome(hwnd) Then
        strUrl = GetChromeDomainName(hwnd)

        If Len(strUrl) > 0 Then
            FrmMain.comKey.Text = FilterDomainName(strUrl)
            SetUrlAsKey = 1
        Else
            SetUrlAsKey = 0

        End If

    ElseIf isFirefox(hwnd) Then
        strUrl = GetFirefoxDomainName(hwnd)

        If Len(strUrl) > 0 Then
            FrmMain.comKey.Text = FilterDomainName(strUrl)
            SetUrlAsKey = 1
        Else
            SetUrlAsKey = 0

        End If

    ElseIf isOpera(hwnd) Then
        strUrl = GetOperaDomainName(hwnd)

        If Len(strUrl) > 0 Then
            FrmMain.comKey.Text = FilterDomainName(strUrl)
            SetUrlAsKey = 1
        Else
            SetUrlAsKey = 0

        End If

    ElseIf isMaxthon(hwnd) Then
        strUrl = GetMaxthonDomainName(hwnd)

        If Len(strUrl) > 0 Then
            FrmMain.comKey.Text = FilterDomainName(strUrl)
            SetUrlAsKey = 1
        Else
            SetUrlAsKey = 0

        End If

    Else
        strUrl = isClipboardAsUrl

        If Len(strUrl) > 0 Then
            FrmMain.comKey.Text = FilterDomainName(strUrl)
            SetUrlAsKey = 1
        Else
            SetUrlAsKey = 0

        End If

    End If

End Function
