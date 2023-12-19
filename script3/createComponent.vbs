Dim catalog
Dim applications
Dim application
Dim components
Dim newComponent

' Set the name of the existing COM+ Application and the path to the DLL
Dim comAppName
Dim dllPath
comAppName = "masivo" ' Replace with the name of your existing COM+ Application
dllPath = "C:\Users\Administrator\Desktop\masivo\ComPolCompag_tx.dll" ' Replace with the full path to your DLL

' Create a new COMAdminCatalog object
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

' Get the collection of COM+ Applications
Set applications = catalog.GetCollection("Applications")
applications.Populate

' Find the existing COM+ Application
Dim appExists
appExists = False
For Each application in applications
    If (application.Name = comAppName) Then
        appExists = True
        Exit For
    End If
Next

If appExists Then
    ' Get the collection of components for the application
    Set components = applications.GetCollection("Components", application.Key)
    components.Populate

    ' Install the new component (DLL) to the COM+ Application
    Dim installResult
    installResult = catalog.InstallComponent(application.Key, dllPath, "", "")

    If installResult Then
        WScript.Echo "Added new component to COM+ Application: " & dllPath
    Else
        WScript.Echo "Failed to add component to COM+ Application."
    End If
Else
    WScript.Echo "COM+ Application not found: " & comAppName
End If

' Clean up
Set newComponent = Nothing
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
