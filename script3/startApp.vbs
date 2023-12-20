Dim catalog
Dim comAppName
Dim appIsRunning

comAppName = "masivo" 

Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

On Error Resume Next

appIsRunning = catalog.IsApplicationRunning(comAppName)

If Err.Number <> 0 Then
    WScript.Echo "Error checking status of COM+ Application '" & comAppName & "': " & Err.Description
    Err.Clear ' Clear the error
ElseIf appIsRunning Then
    WScript.Echo "COM+ Application '" & comAppName & "' is already running."
Else
    ' Start the COM+ Application because it is not already running
    catalog.StartApplication comAppName
    
    If Err.Number <> 0 Then
        WScript.Echo "Error starting COM+ Application '" & comAppName & "': " & Err.Description
        Err.Clear ' Clear the error
    Else
        WScript.Echo "COM+ Application '" & comAppName & "' started successfully."
    End If
End If

On Error GoTo 0 ' Disable error handling

' Clean up
Set catalog = Nothing
