Dim catalog
Dim applications
Dim application
Dim components
Dim appKey

Dim comAppName
Dim dllPath
comAppName = "masivo"
dllPath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"


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
        appKey = application.Key
        Exit For
    End If
Next

If appExists Then
    ' Install the new component (DLL) to the COM+ Application
    On Error Resume Next ' Enable error handling
    catalog.InstallComponent appKey, dllPath, "", ""

    If Err.Number <> 0 Then
        WScript.Echo "Error: " & Err.Description
        Err.Clear ' Clear the error
    Else
        WScript.Echo "Added new component to COM+ Application: " & dllPath
    End If
    On Error GoTo 0 ' Disable error handling
Else
    WScript.Echo "COM+ Application not found: " & comAppName
End If

' Clean up
Set components = Nothing
Set application = Nothing
Set applications = Nothing
Set catalog = Nothing
