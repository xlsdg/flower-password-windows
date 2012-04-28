Attribute VB_Name = "modAutoComplete"
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

Private blnAuto As Boolean

Public Sub cbBox_Change(cbBox As ComboBox)

    Dim strPart As String, iLoop As Long, iStart As Long, strItem As String

    'don't do if no text or if change was made by autocomplete coding
    If Not blnAuto And cbBox.Text <> "" Then
        'save the selection start point (cursor position)
        iStart = cbBox.SelStart
        'get the part the user has typed (not selected)
        strPart = Left$(cbBox.Text, iStart)

        For iLoop = 0 To cbBox.ListCount - 1
            'compare each item to the part the user has typed,
            '"complete" with the first good match
            strItem = UCase$(cbBox.List(iLoop))

            'If strItem Like UCase$(strPart & "*") And strItem <> UCase$(Combobox1.Text) Then
            If strItem Like UCase$(strPart & "*") And strItem <> UCase$(cbBox.Text) And UCase$(strPart) <> strItem Then
                'partial match but not the whole thing.
                '(if whole thing, nothing to complete!)
                blnAuto = True
                cbBox.SelText = Mid$(cbBox.List(iLoop), iStart + 1) 'add on the new ending
                cbBox.SelStart = iStart   'reset the selection
                cbBox.SelLength = Len(cbBox.Text) - iStart
                blnAuto = False
                Exit For

            End If

        Next iLoop

    End If

End Sub

Public Sub cbBox_KeyDown(cbBox As ComboBox, KeyCode As Integer, Shift As Integer)

    'Unless we watch out for it, backspace or delete will just delete
    'the selected text (the autocomplete part), so we delete it here
    'first so it doesn't interfere with what the user expects
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        cbBox.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        cbBox_LostFocus cbBox
        'the following causes the item to be selected and
        'the cursor placed at the end:
        cbBox.SelStart = Len(cbBox.Text)

        'This would select the whole thing instead:
        'combobox1.SelLength = Len(combobox1.Text)
        'alternatively, you could move the focus to the next control here
    End If

End Sub

Public Sub cbBox_LostFocus(cbBox As ComboBox)

    Dim iLoop As Long

    'Match capitalization if item entered is one on the list
    If cbBox.Text <> "" Then

        For iLoop = 0 To cbBox.ListCount - 1

            If UCase$(cbBox.List(iLoop)) = UCase$(cbBox.Text) Then
                blnAuto = True
                cbBox.Text = cbBox.List(iLoop)
                blnAuto = False
                Exit For

            End If

        Next iLoop

    End If

End Sub
