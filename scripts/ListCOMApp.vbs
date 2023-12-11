' Create an instance of the FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a text file to write the output
Set outputFile = fso.CreateTextFile("COMList.txt", True)

' Create an instance of the COM+ Admin catalog
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

' Get the collection of COM+ Applications
Set apps = catalog.GetCollection("Applications")
apps.Populate

' Loop through each application in the collection
For Each app In apps
    ' Write the name of the application to the file
    outputFile.WriteLine(app.Name)
Next

' Close the file
outputFile.Close

' Clean up
Set outputFile = Nothing
Set fso = Nothing
Set apps = Nothing
Set catalog = Nothing
