Sub TestNavigationDirectly()
    ' Test if we can call the navigation function directly
    Dim snipId As String
    snipId = "f0df5445-2c49-4f18-83a6-c765b46e8c82"
    
    ' Try to create the add-in object and call navigation
    On Error GoTo ErrorHandler
    
    Dim addIn As Object
    Set addIn = CreateObject("SnipperPro.Connect")
    
    If Not addIn Is Nothing Then
        ' Call the navigation method directly
        Debug.Print "Calling NavigateToSnip with ID: " & snipId
        ' This would need to be implemented...
        Debug.Print "Add-in object created successfully"
    Else
        Debug.Print "Failed to create add-in object"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
End Sub

Sub TestCellFormula()
    ' Test setting a proper validation formula
    Dim cell As Range
    Set cell = Range("E2")
    
    ' Set the validation formula
    cell.Formula = "=SnipperPro.Connect.VALIDATION(""f0df5445-2c49-4f18-83a6-c765b46e8c82"")"
    
    ' Check what we got
    Debug.Print "Formula set: " & cell.Formula
    Debug.Print "Value shows: " & cell.Value
End Sub 