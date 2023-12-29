Dim catalog
Dim comAppName

Set fso = CreateObject("Scripting.FileSystemObject")
Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\2-stopApp.txt"
Set outputFile = fso.CreateTextFile(outputFilePath, True)

comAppName = "masivo"
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")
On Error Resume Next

catalog.ShutdownApplication comAppName

If Err.Number <> 0 Then
    outputFile.WriteLine("Error stopping COM+ Application '" & comAppName & "': " & Err.Description)
    ' WScript.Echo "Error stopping COM+ Application '" & comAppName & "': " & Err.Description
    Err.Clear
Else
    outputFile.WriteLine("COM+ Application '" & comAppName & "' stopped successfully.")
    ' WScript.Echo "COM+ Application '" & comAppName & "' stopped successfully."
End If

On Error GoTo 0

' Clean up
Set catalog = Nothing
