Attribute VB_Name = "modExplorer"
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

Public Function FilterDomainName(ByVal strDomain As String) As String
    'Dim strExt As String
    'strExt = ".com.cn|.net.cn|.gov.cn|.edu.cn|.org.cn|.mil.cn|.com.hk|.travel|.ac.cn|.bj.cn|.sh.cn|.tj.cn|.cq.cn|.he.cn|.sx.cn|.nm.cn|.ln.cn|.jl.cn|.hl.cn|.js.cn|.zj.cn|.ah.cn|.fj.cn|.jx.cn|.sd.cn|.ha.cn|.hb.cn|.hn.cn|.gd.cn|.gx.cn|.hi.cn|.sc.cn|.gz.cn|.yn.cn|.xz.cn|.sn.cn|.gs.cn|.qh.cn|.nx.cn|.xj.cn|.tw.cn|.hk.cn|.mo.cn|.info|.mobi|.name|.asia|" & _
     ".biz|.cat|.com|.edu|.gov|.int|.mil|.net|.org|.pro|.tel|.xxx|.ac|.ad|.ae|.af|.ag|.ai|.al|.am|.an|.ao|.aq|.as|.at|.aw|.ax|.az|.ba|.bb|.be|.bf|.bg|.bh|.bi|.bj|.bm|.bo|.br|.bs|.bt|.bw|.by|.bz|.ca|.cc|.cd|.cf|.cg|.ch|.ci|.cl|.cm|.cn|.co|.cr|.cu|.cv|.cx|.cz|.de|.dj|.dk|.dm|.do|.dz|.ec|.ee|.es|.eu|.fi|.fm|.fo|.fr|.ga|.gd|.ge|.gf|.gg|.gh|.gi|.gl|.gm|.gp|.gq|.gr|" & _
     ".gs|.gw|.gy|.hk|.hm|.hn|.hr|.ht|.hu|.id|.ie|.im|.in|.io|.iq|.ir|.is|.it|.je|.jo|.jp|.kg|.ki|.km|.kn|.kr|.ky|.kz|.la|.lc|.li|.lk|.ls|.lt|.lu|.lv|.ly|.ma|.mc|.md|.me|.mg|.mh|.mk|.ml|.mn|.mo|.mp|.mq|.mr|.ms|.mu|.mv|.mw|.mx|.my|.na|.nc|.ne|.nf|.nl|.no|.nr|.nu|.pa|.pe|.pf|.ph|.pk|.pl|.pn|.pr|.ps|.pt|.pw|.re|.ro|.rs|.ru|.rw|.sa|.sb|.sc|.sd|.se|.sg|.sh|.si|.sk|" & _
     ".sl|.sm|.sn|.so|.sr|.st|.su|.sy|.sz|.tc|.td|.tf|.tg|.th|.tj|.tk|.tl|.tm|.tn|.to|.tt|.tv|.tw|.ua|.ug|.us|.uz|.va|.vc|.vg|.vi|.vn|.vu|.ws"

    'strExt = strDomains
    Dim arrExt() As String

    arrExt = Split(strDomains, "|")
    strDomain = LCase$(strDomain)

    Dim X As Long

    FilterDomainName = vbNullString

    For X = LBound(arrExt) To UBound(arrExt)

        Dim lenExt As Long, lenStr As Long

        lenExt = Len(arrExt(X))
        lenStr = Len(strDomain)

        If Right$(strDomain, lenExt) = arrExt(X) And lenStr > lenExt Then
            strDomain = Left$(strDomain, lenStr - lenExt)
            lenStr = Len(strDomain)

            Dim v As Long

            v = InStrRev(strDomain, ".")

            If v = 0 Then
                FilterDomainName = strDomain
            Else
                FilterDomainName = Right$(strDomain, lenStr - v)

            End If

            If isDomainSuffix Then '是否包含后缀
                FilterDomainName = FilterDomainName + arrExt(X)

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

    If isAutoCheck And Clipboard.GetFormat(vbCFText) Then '增加自动检测剪贴板开关

        Dim str_url As String, str_len As Long

        str_url = LCase$(Clipboard.GetText)
        str_len = Len(str_url)

        If str_len > 0 Then
            isClipboardAsUrl = vbNullString

            Dim Str_Sites As String

            Str_Sites = LCase$("http|https|ftp|mms|rtsp|rtmp|mmst|bt|www.|ftp.|pop.|smtp.|wap.|m.|3g.")

            Dim arr_ext() As String

            arr_ext = Split(Str_Sites, "|")

            Dim X As Integer

            For X = LBound(arr_ext) To UBound(arr_ext)

                Dim arr_len As Long

                arr_len = Len(arr_ext(X))

                If Left$(str_url, arr_len) = arr_ext(X) And str_len > arr_len Then
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
    ElseIf isChrome(hwnd) Then
        strUrl = GetChromeDomainName(hwnd)
    ElseIf isFirefox(hwnd) Then
        strUrl = GetFirefoxDomainName(hwnd)
    ElseIf isOpera(hwnd) Then
        strUrl = GetOperaDomainName(hwnd)
    ElseIf isMaxthon(hwnd) Then
        strUrl = GetMaxthonDomainName(hwnd)
    Else
        strUrl = isClipboardAsUrl

    End If

    If Len(strUrl) > 0 Then
        strUrl = FilterDomainName(strUrl)

        If Len(strUrl) > 0 Then
            FrmMain.comKey.Text = strUrl
            SetUrlAsKey = 1
        Else
            SetUrlAsKey = 0

        End If

    Else
        SetUrlAsKey = 0

    End If

End Function
