Dim catalog
Dim comAppName

comAppName = "masivo"

Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

On Error Resume Next

catalog.StartApplication comAppName

If Err.Number <> 0 Then
    If InStr(Err.Description, "The application is already running") > 0 Then
        WScript.Echo "COM+ Application '" & comAppName & "' is already running."
    Else
        WScript.Echo "Error starting COM+ Application '" & comAppName & "': " & Err.Description
    End If
    Err.Clear
Else
    WScript.Echo "COM+ Application '" & comAppName & "' started successfully."
End If

On Error GoTo 0

' Clean up
Set catalog = Nothing
