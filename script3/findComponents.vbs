Dim catalog
Dim applications
Dim application
Dim components
Dim component
Dim fso, outputFile

' Specify the output file path
Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\ComponentsList.txt"

' Create a new COMAdminCatalog object
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

' Get the collection of COM+ Applications
Set applications = catalog.GetCollection("Applications")

applications.Populate

' Create FileSystemObject to handle file writing
Set fso = CreateObject("Scripting.FileSystemObject")
Set outputFile = fso.CreateTextFile(outputFilePath, True)

' Loop through the applications
For Each application in applications
    If (application.Name = "masivo") Then
        ' Get the collection of components for the application
        Set components = applications.GetCollection("Components", application.Key)
        components.Populate

        ' Write all components to file
        For Each component in components
            outputFile.WriteLine("Component: " & component.Name)
        Next
    End If
Next

' Clean up
outputFile.Close
Set outputFile = Nothing
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
Set fso = Nothing
