Dim catalog
Dim comAppName

' Set the name of the existing COM+ Application
comAppName = "masivo" ' Replace with the name of your existing COM+ Application

' Create a new COMAdminCatalog object
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

On Error Resume Next ' Enable error handling

' Attempt to start the COM+ Application
catalog.StartApplication comAppName

If Err.Number <> 0 Then
    ' Check if the error is because the application is already running or another issue
    If InStr(Err.Description, "The application is already running") > 0 Then
        WScript.Echo "COM+ Application '" & comAppName & "' is already running."
    Else
        WScript.Echo "Error starting COM+ Application '" & comAppName & "': " & Err.Description
    End If
    Err.Clear ' Clear the error
Else
    WScript.Echo "COM+ Application '" & comAppName & "' started successfully."
End If

On Error GoTo 0 ' Disable error handling

' Clean up
Set catalog = Nothing
