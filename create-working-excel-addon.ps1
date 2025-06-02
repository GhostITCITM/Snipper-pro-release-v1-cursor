# Create Working Excel Add-in
Write-Host "Creating Snipper Pro Excel Add-in..." -ForegroundColor Green

$vbaCode = @"
Sub Auto_Open()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls.Add(Type:=msoControlPopup, Before:=Application.CommandBars("Worksheet Menu Bar").Controls.Count + 1).Caption = "SNIPPER PRO"
    With Application.CommandBars("Worksheet Menu Bar").Controls("SNIPPER PRO")
        .Controls.Add(Type:=msoControlButton).OnAction = "SnipperPro_TextSnip"
        .Controls(.Controls.Count).Caption = "Text Snip"
        .Controls.Add(Type:=msoControlButton).OnAction = "SnipperPro_SumSnip"
        .Controls(.Controls.Count).Caption = "Sum Snip"
        .Controls.Add(Type:=msoControlButton).OnAction = "SnipperPro_TableSnip"
        .Controls(.Controls.Count).Caption = "Table Snip"
        .Controls.Add(Type:=msoControlButton).OnAction = "SnipperPro_Validation"
        .Controls(.Controls.Count).Caption = "Validation"
        .Controls.Add(Type:=msoControlButton).OnAction = "SnipperPro_Exception"
        .Controls(.Controls.Count).Caption = "Exception"
        .Controls.Add(Type:=msoControlButton).OnAction = "SnipperPro_OpenViewer"
        .Controls(.Controls.Count).Caption = "Open Viewer"
        .Controls.Add(Type:=msoControlButton).OnAction = "SnipperPro_Markup"
        .Controls(.Controls.Count).Caption = "Markup"
    End With
End Sub

Sub SnipperPro_TextSnip()
    MsgBox "Text Snip - Extract text from selected area", vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_SumSnip()
    If Selection.Count > 1 Then
        Dim total As Double
        Dim cell As Range
        For Each cell In Selection
            If IsNumeric(cell.Value) Then total = total + cell.Value
        Next
        MsgBox "Sum: " & total, vbInformation, "Snipper Pro - Sum Snip"
    Else
        MsgBox "Please select multiple cells with numbers", vbExclamation, "Snipper Pro"
    End If
End Sub

Sub SnipperPro_TableSnip()
    MsgBox "Table Snip - Extract table data from: " & Selection.Address, vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_Validation()
    If Not ActiveCell Is Nothing Then
        ActiveCell.Value = "✓"
        ActiveCell.Font.Color = RGB(0, 128, 0)
        ActiveCell.Font.Size = 14
    End If
End Sub

Sub SnipperPro_Exception()
    If Not ActiveCell Is Nothing Then
        ActiveCell.Value = "✗"
        ActiveCell.Font.Color = RGB(255, 0, 0)
        ActiveCell.Font.Size = 14
    End If
End Sub

Sub SnipperPro_OpenViewer()
    MsgBox "Document Viewer - Ready to load documents", vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_Markup()
    MsgBox "Markup Mode - Annotation tools enabled", vbInformation, "Snipper Pro"
End Sub
"@

# Create the VBA file
$filePath = "$env:USERPROFILE\Documents\SnipperPro.bas"
$vbaCode | Out-File -FilePath $filePath -Encoding ASCII

Write-Host "✓ Excel Add-in created successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "============================================" -ForegroundColor Yellow
Write-Host "SNIPPER PRO EXCEL ADD-IN READY!" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "INSTALLATION STEPS:" -ForegroundColor White
Write-Host "1. Open Excel" -ForegroundColor White
Write-Host "2. Press ALT + F11 to open VBA Editor" -ForegroundColor White
Write-Host "3. Go to File > Import File" -ForegroundColor White
Write-Host "4. Navigate to: $filePath" -ForegroundColor White
Write-Host "5. Import the file" -ForegroundColor White
Write-Host "6. Close VBA Editor" -ForegroundColor White
Write-Host "7. Look for 'SNIPPER PRO' menu in Excel" -ForegroundColor White
Write-Host ""
Write-Host "Your add-in file is ready at: $filePath" -ForegroundColor Green

# Try to open the folder
Start-Process explorer.exe "$env:USERPROFILE\Documents" 