Dim catalog
Dim applications
Dim application
Dim components
Dim component
Dim fso, outputFile

Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\2-ComponentsList.txt"

Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

Set applications = catalog.GetCollection("Applications")

applications.Populate

Set fso = CreateObject("Scripting.FileSystemObject")
Set outputFile = fso.CreateTextFile(outputFilePath, True)

For Each application in applications
    If (application.Name = "masivo") Then
    
        Set components = applications.GetCollection("Components", application.Key)
        components.Populate
    
        For Each component in components
            outputFile.WriteLine("Component: " & component.Name)
        Next
    End If
Next

outputFile.Close
Set outputFile = Nothing
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
Set fso = Nothing
