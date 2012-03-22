Attribute VB_Name = "modExplorer"
Option Explicit

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
        SetUrlAsKey = 0

    End If

End Function

Public Function FilterDomainName(ByVal Str_DomainName As String) As String

    Dim Str_Extensions As String

    Str_Extensions = ".com.cn|.net.cn|.gov.cn|.edu.cn|.org.cn|.mil.cn|.ac.cn|.bj.cn|.sh.cn|.tj.cn|.cq.cn|.he.cn|.sx.cn|.nm.cn|.ln.cn|.jl.cn|.hl.cn|.js.cn|.zj.cn|.ah.cn|.fj.cn|.jx.cn|.sd.cn|.ha.cn|.hb.cn|.hn.cn|.gd.cn|.gx.cn|.hi.cn|.sc.cn|.gz.cn|.yn.cn|.xz.cn|.sn.cn|.gs.cn|.qh.cn|.nx.cn|.xj.cn|.tw.cn|.hk.cn|.mo.cn|.com.hk|.com|.net|.org|.hk|.cc|.info|.biz|.me|.us|.cc|.info|.mobi|.name|.asia|.travel|.tel|.tv|.mil|.int|.edu"

    Dim Arr_Ext() As String

    Arr_Ext = Split(Str_Extensions, "|")
    Str_DomainName = LCase$(Str_DomainName)

    Dim x As Integer

    For x = LBound(Arr_Ext) To UBound(Arr_Ext)

        Dim Arr_len As Long, Str_len As Long

        Arr_len = Len(Arr_Ext(x))
        Str_len = Len(Str_DomainName)

        If Right$(Str_DomainName, Arr_len) = Arr_Ext(x) And Str_len > Arr_len Then
            Str_DomainName = Left$(Str_DomainName, Str_len - Arr_len)
            Str_len = Len(Str_DomainName)

            Dim v As Integer

            v = InStrRev(Str_DomainName, ".")

            If v = 0 Then
                FilterDomainName = Str_DomainName
            Else
                FilterDomainName = Right$(Str_DomainName, Str_len - v)

            End If

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
