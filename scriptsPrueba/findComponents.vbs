Dim catalog
Dim applications
Dim application
Dim components
Dim component

' Create a new COMAdminCatalog object
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

' Get the collection of COM+ Applications
Set applications = catalog.GetCollection("Applications")

applications.Populate

' Loop through the applications
For Each application in applications
    If (application.Name = "masivo") Then
        ' Get the collection of components for the application
        Set components = applications.GetCollection("Components", application.Key)
        components.Populate

        ' Loop through the components to find your DLL
        For Each component in components
            If (InStr(component.Name, "YourDLLName") > 0) Then ' Replace with your DLL name
                WScript.Echo "Found DLL: " & component.Name
            End If
        Next
    End If
Next

' Clean up
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
