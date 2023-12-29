Dim catalog
Dim applications
Dim application
Dim components
Dim appKey
Set fso = CreateObject("Scripting.FileSystemObject")

Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\4-create.txt"
Set outputFile = fso.CreateTextFile(outputFilePath, True)

Dim comAppName
Dim dllPath
comAppName = "masivo"
dllPath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"

If fso.FileExists(dllPath) Then
    outputFile.WriteLine("DLL file found: " & dllPath)
    ' WScript.Echo "DLL file found: " & dllPath
Else
    outputFile.WriteLine("DLL file not found: " & dllPath)
    ' WScript.Echo "Error: DLL file NOT found: " & dllPath
    WScript.Quit(1)
End If

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
        outputFile.WriteLine("Error: " & Err)
        outputFile.WriteLine("Error: Error al crear nuevo componente, favor validar dll generada")
        ' WScript.Echo "Error: " & Err
        Err.Clear
    Else
        outputFile.WriteLine("Added new component to COM+ Application: " & dllPath)
        ' WScript.Echo "Added new component to COM+ Application: " & dllPath
        WScript.Quit(0)
    End If
    On Error GoTo 0
Else
    outputFile.WriteLine("COM+ Application not found:" & comAppName)
    ' WScript.Echo "COM+ Application not found: " & comAppName
    WScript.Quit(1) ' Exit script with an error code
End If

Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing

