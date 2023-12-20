Dim catalog
Dim applications
Dim application
Dim components
Dim appKey
Set fso = CreateObject("Scripting.FileSystemObject")

Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\create.txt"
Set outputFile = fso.CreateTextFile(outputFilePath, True)

Dim comAppName
Dim dllPath
comAppName = "masivo"
dllPath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"

Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

Set applications = catalog.GetCollection("Applications")
applications.Populate

Dim appExists
appExists = False
For Each application in applications
    If (application.Name = comAppName) Then
        appExists = True
        appKey = application.Key
        Exit For
    End If
Next

If appExists Then

    On Error Resume Next
    catalog.InstallComponent appKey, dllPath, "", ""

    If Err.Number <> 0 Then
        outputFile.WriteLine("Error: " & Err.Description)
        WScript.Echo "Error: " & Err.Description
        Err.Clear
    Else
        outputFile.WriteLine("Added new component to COM+ Application: " & dllPath)
        WScript.Echo "Added new component to COM+ Application: " & dllPath
    End If
    On Error GoTo 0
Else
    outputFile.WriteLine("COM+ Application not found:" & comAppName)
    WScript.Echo "COM+ Application not found: " & comAppName
End If

Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
