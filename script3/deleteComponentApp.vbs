Dim catalog
Dim applications
Dim application
Dim components
Dim component
Dim componentNameToDelete

Set fso = CreateObject("Scripting.FileSystemObject")
Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\Delete.txt"
Set outputFile = fso.CreateTextFile(outputFilePath, True)

Dim comAppName
comAppName = "masivo"
componentNameToDelete = "comPOLCOMPAG_TX.impCompag"

Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

Set applications = catalog.GetCollection("Applications")

applications.Populate

For Each application in applications
    If (application.Name = comAppName) Then
        Set components = applications.GetCollection("Components", application.Key)
        components.Populate

        Dim i
        For i = 0 To components.Count - 1
            Set component = components.Item(i)
            If (component.Name = componentNameToDelete) Then
                ' Delete the component
                components.Remove(i)
                ' Save the changes to the COM+ Catalog
                components.SaveChanges
                outputFile.WriteLine("Component deleted: " & componentNameToDelete)
                WScript.Echo "Component deleted: " & componentNameToDelete
                Exit For
            End If
        Next
    End If
Next

outputFile.Close

Set outputFile = Nothing
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
