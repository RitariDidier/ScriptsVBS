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
    catalog.InstallComponent application.Key, dllPath, "", ""
    
    ' Check for errors
    If catalog.GetCollection("Errors").Count > 0 Then
        Dim errorMessages, item
        Set errorMessages = catalog.GetCollection("Errors")
        errorMessages.Populate

        ' Display each error message
        For Each item in errorMessages
            WScript.Echo "Error: " & item.Value("Description")
        Next
    Else
        WScript.Echo "Added new component to COM+ Application: " & dllPath
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
