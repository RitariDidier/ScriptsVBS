Dim catalog
Dim applications
Dim application
Dim components
Dim component
Dim componentNameToDelete
Dim found

Set fso = CreateObject("Scripting.FileSystemObject")
Dim outputFilePath
outputFilePath = "C:\Users\Administrator\Desktop\logs\3-deleteComponentApp.txt"
Set outputFile = fso.CreateTextFile(outputFilePath, True)

Dim comAppName
comAppName = "masivo"
componentNameToDelete = "comPOLCOMPAG_TX.impCompag"

Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

Set applications = catalog.GetCollection("Applications")

applications.Populate

found = False

' validate existence of component
For Each application in applications
    If (application.Name = comAppName) Then
        WScript.Echo "ComApp Found: "  & comAppName
        Set components = applications.GetCollection("Components", application.Key)
        components.Populate

        Dim i
        For i = 0 To components.Count - 1
            Set component = components.Item(i)
            If (component.Name = componentNameToDelete) Then
                WScript.Echo "Component Found: " & componentNameToDelete
                found = True
                Exit For
            End If
        Next
    End If
Next

if (Found) Then
    For Each application in applications
        If (application.Name = comAppName) Then
            Set components = applications.GetCollection("Components", application.Key)
            components.Populate

            ' Dim i
            For i = 0 To components.Count - 1
                Set component = components.Item(i)
                If (component.Name = componentNameToDelete) Then
                    components.Remove(i)

                    components.SaveChanges
                    outputFile.WriteLine("Component deleted: " & componentNameToDelete)
                    WScript.Echo "Component deleted: " & componentNameToDelete
                    Exit For
                End If
            Next
        End If
    Next
else
    outputFile.WriteLine("Error: Component Not Found: " & componentNameToDelete)
    WScript.Echo("Error: Component Not Found: " & componentNameToDelete)
    WScript.Quit(1) ' Exit script with an error code
End If



outputFile.Close

Set outputFile = Nothing
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
