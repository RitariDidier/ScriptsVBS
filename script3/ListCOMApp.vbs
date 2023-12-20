
Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\COMList.txt"

Set fso = CreateObject("Scripting.FileSystemObject")

Set outputFile = fso.CreateTextFile(outputFilePath, True)

Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

Set apps = catalog.GetCollection("Applications")
apps.Populate

For Each app In apps
    outputFile.WriteLine(app.Name)
Next

outputFile.Close

Set outputFile = Nothing
Set fso = Nothing
Set apps = Nothing
Set catalog = Nothing
