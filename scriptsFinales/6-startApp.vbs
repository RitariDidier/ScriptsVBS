Dim catalog
Dim comAppName


Set fso = CreateObject("Scripting.FileSystemObject")
Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\5-startComApp.txt"
Set outputFile = fso.CreateTextFile(outputFilePath, True)

comAppName = "masivo"
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")
On Error Resume Next

catalog.StartApplication comAppName

If Err.Number <> 0 Then
    If InStr(Err.Description, "The application is already running") > 0 Then
        outputFile.WriteLine("COM+ Application '" & comAppName & "' is already running.")
        ' WScript.Echo "COM+ Application '" & comAppName & "' is already running."
    Else
        outputFile.WriteLine("Error starting COM+ Application '" & comAppName & "': " & Err.Description)
        ' WScript.Echo "Error starting COM+ Application '" & comAppName & "': " & Err.Description
    End If
    Err.Clear
Else
    outputFile.WriteLine("COM+ Application '" & comAppName & "' started successfully.")
    ' WScript.Echo "COM+ Application '" & comAppName & "' started successfully."
    WScript.Quit(0)
End If

On Error GoTo 0

' Clean up
Set catalog = Nothing
