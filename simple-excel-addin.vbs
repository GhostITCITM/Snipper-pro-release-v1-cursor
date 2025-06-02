Sub SnipperPro_TextSnip()
    MsgBox "Text Snip - Extract text from selected area", vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_SumSnip()
    MsgBox "Sum Snip - Sum numbers from selected area", vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_TableSnip()
    MsgBox "Table Snip - Extract table data", vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_Validation()
    If Not ActiveCell Is Nothing Then
        ActiveCell.Value = "✓"
    End If
End Sub

Sub SnipperPro_Exception()
    If Not ActiveCell Is Nothing Then
        ActiveCell.Value = "✗"
    End If
End Sub

Sub SnipperPro_OpenViewer()
    MsgBox "Document Viewer - Open document viewer", vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_Markup()
    MsgBox "Markup - Toggle markup mode", vbInformation, "Snipper Pro"
End Sub 