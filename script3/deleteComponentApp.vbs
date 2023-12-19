Dim catalog
Dim applications
Dim application
Dim components
Dim component
Dim componentNameToDelete

' Set the name of the COM+ Application and the component to delete
Dim comAppName
comAppName = "masivo" ' Replace with your COM+ Application name
componentNameToDelete = "comPOLCOMPAG_TX.impCompag" ' Replace with the name of the component to delete

' Create a new COMAdminCatalog object
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

' Get the collection of COM+ Applications
Set applications = catalog.GetCollection("Applications")

applications.Populate

' Loop through the applications
For Each application in applications
    If (application.Name = comAppName) Then
        ' Get the collection of components for the application
        Set components = applications.GetCollection("Components", application.Key)
        components.Populate

        ' Loop through the components to find the one to delete
        Dim i
        For i = 0 To components.Count - 1
            Set component = components.Item(i)
            If (component.Name = componentNameToDelete) Then
                ' Delete the component
                components.Remove(i)
                ' Save the changes to the COM+ Catalog
                components.SaveChanges
                WScript.Echo "Component deleted: " & componentNameToDelete
                Exit For
            End If
        Next
    End If
Next

' Clean up
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
