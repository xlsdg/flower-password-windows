Attribute VB_Name = "modRoundFrm"
Option Explicit

Private Declare Function SetWindowRgn _
                Lib "user32.dll" (ByVal hwnd As Long, _
                                  ByVal hRgn As Long, _
                                  ByVal bRedraw As Boolean) As Long

Private Declare Function CreateRoundRectRgn _
                Lib "gdi32.dll" (ByVal X1 As Long, _
                                 ByVal Y1 As Long, _
                                 ByVal X2 As Long, _
                                 ByVal Y2 As Long, _
                                 ByVal X3 As Long, _
                                 ByVal Y3 As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private outrgn As Long

Public Function SetFormRgn() As Long
    SetFormRgn = rgnForm(FrmMain, 15, 15) '调用子过程

End Function

Public Function UnSetFormRgn() As Long
    UnSetFormRgn = DeleteObject(outrgn) '将圆角区域使用的所有系统资源释放

End Function

Private Function rgnForm(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) As Long

    Dim w As Long, h As Long, outrgn As Long

    w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
    h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
    outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
    rgnForm = SetWindowRgn(frmbox.hwnd, outrgn, True)

End Function
