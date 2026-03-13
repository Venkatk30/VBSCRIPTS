Function SqlInjectCheck(args)
    ' ==========================
    ' VBScript: Dynamic Parameter Validation with Regex
    ' ==========================
    
    On Error Resume Next
    
    ' Regex to allow only letters, numbers, @ . _ - and spaces
    ' Disallow common SQL injection characters: ' " ; --
    Dim regEx
    Set regEx = New RegExp
    regEx.Pattern = "^[a-zA-Z0-9@._ -]+$"
    regEx.IgnoreCase = True
    
    ' Loop through all parameters (array)
    For i = LBound(args) To UBound(args)
        paramValue = args(i)
        
        ' Check if numeric
        If IsNumeric(paramValue) Then
            ' Valid number
        Else
            ' Validate allowed characters
            If Not regEx.Test(paramValue) Then
                SqlInjectCheck = "Invalid parameter at position " & (i+1) & ": contains forbidden characters."
                Exit Function
            End If
            
            ' Additional check for SQL keywords
            Dim badWords, w
            badWords = Array("DROP", "DELETE", "INSERT", "UPDATE", "SELECT", "OR", "XP_", "--")
            
            For Each w In badWords
                If InStr(UCase(paramValue), w) > 0 Then
                    SqlInjectCheck = "Invalid parameter at position " & (i+1) & ": contains SQL keywords or patterns (" & w & ")"
                    Exit Function
                End If
            Next
        End If
    Next
    
    If Err.Number <> 0 Then
        SqlInjectCheck = "Error Number: " & Err.Number & " Description: " & Err.Description
        Err.Clear
    Else
        SqlInjectCheck = "All parameters are valid."
    End If
    
    On Error GoTo 0
End Function