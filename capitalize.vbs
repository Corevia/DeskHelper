' By CoreVia 2024
' Redistribution or sharing of this software, in whole or in part, is prohibited.

Sub CapitalizeFirstLetter()
    Dim selectedText As Range
    Dim word As Range

    ' Check if text is selected
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select some text first.", vbExclamation
        Exit Sub
    End If

    ' Loop through each word in the selected range
    Set selectedText = Selection.Range
    For Each word In selectedText.Words
        ' Trim any leading or trailing spaces
        word.Text = Trim(word.Text)

        ' Capitalize the first letter and make the rest lowercase
        word.Text = UCase(Left(word.Text, 1)) & LCase(Mid(word.Text, 2))
    Next word
End Sub
