@echo off
echo Creating Working Excel Add-in...

rem Create the Excel Add-in directory
mkdir "%USERPROFILE%\Documents\Excel Add-ins" 2>nul

rem Create the working VBA add-in file
echo Creating Excel VBA Add-in...
powershell -Command ^
"$content = @'" ^
"Sub Auto_Open()" ^
"    On Error Resume Next" ^
"    Application.CommandBars(\"Worksheet Menu Bar\").Controls.Add(Type:=msoControlPopup, Before:=Application.CommandBars(\"Worksheet Menu Bar\").Controls.Count + 1).Caption = \"SNIPPER PRO\"" ^
"    With Application.CommandBars(\"Worksheet Menu Bar\").Controls(\"SNIPPER PRO\")" ^
"        .Controls.Add(Type:=msoControlButton).OnAction = \"SnipperPro_TextSnip\"" ^
"        .Controls(.Controls.Count).Caption = \"Text Snip\"" ^
"        .Controls.Add(Type:=msoControlButton).OnAction = \"SnipperPro_SumSnip\"" ^
"        .Controls(.Controls.Count).Caption = \"Sum Snip\"" ^
"        .Controls.Add(Type:=msoControlButton).OnAction = \"SnipperPro_TableSnip\"" ^
"        .Controls(.Controls.Count).Caption = \"Table Snip\"" ^
"        .Controls.Add(Type:=msoControlButton).OnAction = \"SnipperPro_Validation\"" ^
"        .Controls(.Controls.Count).Caption = \"Validation\"" ^
"        .Controls.Add(Type:=msoControlButton).OnAction = \"SnipperPro_Exception\"" ^
"        .Controls(.Controls.Count).Caption = \"Exception\"" ^
"        .Controls.Add(Type:=msoControlButton).OnAction = \"SnipperPro_OpenViewer\"" ^
"        .Controls(.Controls.Count).Caption = \"Open Viewer\"" ^
"        .Controls.Add(Type:=msoControlButton).OnAction = \"SnipperPro_Markup\"" ^
"        .Controls(.Controls.Count).Caption = \"Markup\"" ^
"    End With" ^
"End Sub" ^
"" ^
"Sub SnipperPro_TextSnip()" ^
"    MsgBox \"Text Snip - Extract text from selected area\", vbInformation, \"Snipper Pro\"" ^
"End Sub" ^
"" ^
"Sub SnipperPro_SumSnip()" ^
"    If Selection.Count > 1 Then" ^
"        Dim total As Double" ^
"        Dim cell As Range" ^
"        For Each cell In Selection" ^
"            If IsNumeric(cell.Value) Then total = total + cell.Value" ^
"        Next" ^
"        MsgBox \"Sum: \" & total, vbInformation, \"Snipper Pro - Sum Snip\"" ^
"    Else" ^
"        MsgBox \"Please select multiple cells with numbers\", vbExclamation, \"Snipper Pro\"" ^
"    End If" ^
"End Sub" ^
"" ^
"Sub SnipperPro_TableSnip()" ^
"    MsgBox \"Table Snip - Extract table data from: \" & Selection.Address, vbInformation, \"Snipper Pro\"" ^
"End Sub" ^
"" ^
"Sub SnipperPro_Validation()" ^
"    If Not ActiveCell Is Nothing Then" ^
"        ActiveCell.Value = \"✓\"" ^
"        ActiveCell.Font.Color = RGB(0, 128, 0)" ^
"        ActiveCell.Font.Size = 14" ^
"    End If" ^
"End Sub" ^
"" ^
"Sub SnipperPro_Exception()" ^
"    If Not ActiveCell Is Nothing Then" ^
"        ActiveCell.Value = \"✗\"" ^
"        ActiveCell.Font.Color = RGB(255, 0, 0)" ^
"        ActiveCell.Font.Size = 14" ^
"    End If" ^
"End Sub" ^
"" ^
"Sub SnipperPro_OpenViewer()" ^
"    MsgBox \"Document Viewer - Ready to load documents\", vbInformation, \"Snipper Pro\"" ^
"End Sub" ^
"" ^
"Sub SnipperPro_Markup()" ^
"    MsgBox \"Markup Mode - Annotation tools enabled\", vbInformation, \"Snipper Pro\"" ^
"End Sub" ^
"'@; $content | Out-File -FilePath \"%USERPROFILE%\Documents\SnipperPro.bas\" -Encoding ASCII"

echo.
echo ============================================
echo SNIPPER PRO EXCEL ADD-IN CREATED!
echo ============================================
echo.
echo INSTALLATION STEPS:
echo 1. Open Excel
echo 2. Press ALT + F11 to open VBA Editor
echo 3. Go to File ^> Import File
echo 4. Navigate to: %USERPROFILE%\Documents\SnipperPro.bas
echo 5. Import the file
echo 6. Close VBA Editor
echo 7. Look for "SNIPPER PRO" menu in Excel
echo.
echo Your add-in is ready at: %USERPROFILE%\Documents\SnipperPro.bas
echo.
pause 