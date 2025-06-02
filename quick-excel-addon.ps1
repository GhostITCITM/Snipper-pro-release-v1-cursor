Write-Host "Creating Excel Add-in..." -ForegroundColor Green

$vbaContent = @"
Sub Auto_Open()
    On Error Resume Next
    Dim menuBar As CommandBar
    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    
    Dim newMenu As CommandBarControl
    Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup)
    newMenu.Caption = "SNIPPER PRO"
    
    With newMenu
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Text Snip"
            .OnAction = "SnipperPro_TextSnip"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Sum Snip"
            .OnAction = "SnipperPro_SumSnip"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Validation"
            .OnAction = "SnipperPro_Validation"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Exception"
            .OnAction = "SnipperPro_Exception"
        End With
    End With
End Sub

Sub SnipperPro_TextSnip()
    MsgBox "Text Snip functionality", vbInformation, "Snipper Pro"
End Sub

Sub SnipperPro_SumSnip()
    If Selection.Count > 1 Then
        Dim total As Double
        total = 0
        Dim cell As Range
        For Each cell In Selection
            If IsNumeric(cell.Value) Then 
                total = total + cell.Value
            End If
        Next cell
        MsgBox "Sum: " & total, vbInformation, "Snipper Pro"
    Else
        MsgBox "Select multiple cells with numbers", vbInformation, "Snipper Pro"
    End If
End Sub

Sub SnipperPro_Validation()
    If Not ActiveCell Is Nothing Then
        ActiveCell.Value = "✓"
        ActiveCell.Font.Color = RGB(0, 128, 0)
    End If
End Sub

Sub SnipperPro_Exception()
    If Not ActiveCell Is Nothing Then
        ActiveCell.Value = "✗"
        ActiveCell.Font.Color = RGB(255, 0, 0)
    End If
End Sub
"@

$outputPath = "$env:USERPROFILE\Documents\SnipperPro.bas"
$vbaContent | Out-File -FilePath $outputPath -Encoding ASCII

Write-Host "SUCCESS! Excel Add-in created!" -ForegroundColor Green
Write-Host "File location: $outputPath" -ForegroundColor Yellow
Write-Host ""
Write-Host "TO INSTALL:" -ForegroundColor White
Write-Host "1. Open Excel" -ForegroundColor White
Write-Host "2. Press ALT+F11" -ForegroundColor White
Write-Host "3. File > Import File" -ForegroundColor White
Write-Host "4. Select: $outputPath" -ForegroundColor White
Write-Host "5. Close VBA Editor and check Excel menu" -ForegroundColor White 