Attribute VB_Name = "modUTF8"
Option Explicit

Private Declare Function WideCharToMultiByte _
                Lib "kernel32.dll" (ByVal CodePage As Long, _
                                    ByVal dwFlags As Long, _
                                    ByVal lpWideCharStr As Long, _
                                    ByVal cchWideChar As Long, _
                                    ByRef lpMultiByteStr As Any, _
                                    ByVal cchMultiByte As Long, _
                                    ByVal lpDefaultChar As String, _
                                    ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 = 65001

Public Function Hash2Byte(ByVal strValue As String) As Byte()

    If Len(strValue) > 0 Then

        Dim bytBuffer(0 To 15) As Byte

        bytBuffer(0) = CByte("&H" + Mid$(strValue, 1, 2))
        bytBuffer(1) = CByte("&H" + Mid$(strValue, 3, 2))
        bytBuffer(2) = CByte("&H" + Mid$(strValue, 5, 2))
        bytBuffer(3) = CByte("&H" + Mid$(strValue, 7, 2))
        bytBuffer(4) = CByte("&H" + Mid$(strValue, 9, 2))
        bytBuffer(5) = CByte("&H" + Mid$(strValue, 11, 2))
        bytBuffer(6) = CByte("&H" + Mid$(strValue, 13, 2))
        bytBuffer(7) = CByte("&H" + Mid$(strValue, 15, 2))
        bytBuffer(8) = CByte("&H" + Mid$(strValue, 17, 2))
        bytBuffer(9) = CByte("&H" + Mid$(strValue, 19, 2))
        bytBuffer(10) = CByte("&H" + Mid$(strValue, 21, 2))
        bytBuffer(11) = CByte("&H" + Mid$(strValue, 23, 2))
        bytBuffer(12) = CByte("&H" + Mid$(strValue, 25, 2))
        bytBuffer(13) = CByte("&H" + Mid$(strValue, 27, 2))
        bytBuffer(14) = CByte("&H" + Mid$(strValue, 29, 2))
        bytBuffer(15) = CByte("&H" + Mid$(strValue, 31, 2))
        Hash2Byte = bytBuffer

    End If

End Function

Public Function UnicodeToUtf8(ByVal UCS As String) As Byte()

    Dim lLength As Long

    lLength = Len(UCS)

    If lLength > 0 Then

        Dim lBufferSize As Long

        Dim lResult     As Long

        Dim abUTF8()    As Byte

        lBufferSize = lLength * 3 + 1
        ReDim abUTF8(lBufferSize - 1)
        lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, abUTF8(0), lBufferSize, vbNullString, 0)

        If lResult <> 0 Then
            lResult = lResult - 1
            ReDim Preserve abUTF8(lResult)
            UnicodeToUtf8 = abUTF8

        End If

    End If

End Function
