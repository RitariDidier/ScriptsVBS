' Create an instance of the COM+ Admin catalog
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

' Get the collection of COM+ Applications
Set apps = catalog.GetCollection("Applications")
apps.Populate

' The name of the COM+ Application containing the component
Dim applicationName
applicationName = "masivo"

' The name of the component you want to delete
Dim componentName
componentName = "comPOLCOMPAG_TX.impCompag"

Dim appFound
appFound = False

Dim componentFound
componentFound = False

' File to write the result to
Dim resultFile
resultFile = "C:\Users\Administrator\Desktop\logs\Delete.txt"

' Open the file for writing
Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set outputFile = fileSystem.OpenTextFile(resultFile, 2, True)

' Iterate through the applications to find the target application
For Each app In apps
    If app.Name = applicationName Then
        appFound = True

        ' Get the components of the application
        Set components = apps.GetCollection("Components", app.Key)
        components.Populate

        ' Iterate through the components to find the target component
        For Each component In components
            If component.Name = componentName Then
                componentFound = True

                ' Remove the component
                components.Remove component.Key
                components.SaveChanges

                outputFile.WriteLine("Component '" & componentName & "' has been removed from '" & applicationName & "'.")
                Exit For
            End If
        Next

        Exit For
    End If
Next

If Not appFound Then
    outputFile.WriteLine("Application '" & applicationName & "' not found.")
End If

If appFound And Not componentFound Then
    outputFile.WriteLine("Component '" & componentName & "' not found in '" & applicationName & "'.")
End If

' Clean up
outputFile.Close
Set outputFile = Nothing
Set fileSystem = Nothing
Set component = Nothing
Set components = Nothing
Set app = Nothing
Set apps = Nothing
Set catalog = Nothing
