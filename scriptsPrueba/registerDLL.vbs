' |                              |
' ||||||||||||||||||||||||||||||||
' |        Registra DLL          |
' ||||||||||||||||||||||||||||||||
' |                              |
' Path to the DLL you want to register
Dim componentDll
' carpeta del asunto #TODO: DESPUES DE QUE SE COPIA
componentDll = "C:\Components\masivo\ComPolCompag_tx.dll"

' Register the DLL (this requires administrative privileges)
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "regsvr32 /s """ & componentDll & """", 0, True



' |                              |
' ||||||||||||||||||||||||||||||||
' |         FindProgID           |
' ||||||||||||||||||||||||||||||||
' |                              |

' Set the name of the DLL file
Dim dllName
dllName = "ComPolCompag_tx.dll" ' Replace with your DLL file name

' Registry path to start the search
Dim regPath
regPath = "HKEY_CLASSES_ROOT\"

' Command to search the registry
Dim command
command = "reg query " & Chr(34) & regPath & Chr(34) & " /s /f " & Chr(34) & dllName & Chr(34) & " /t REG_SZ"

' Run the command and capture the output
Dim output
output = shell.Exec(command).StdOut.ReadAll

' Display the output
WScript.Echo output

' Clean up
Set shell = Nothing



' |                              |
' ||||||||||||||||||||||||||||||||
' |         FindCLSID            |
' ||||||||||||||||||||||||||||||||
' |                              |

' Set the ProgID of your component
Dim progID
progID = "Component.ProgID"

' Create a Registry object
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' Registry path for ProgID
Dim regPath
regPath = "HKEY_CLASSES_ROOT\" & progID & "\CLSID\"

' Read the CLSID from the registry
Dim clsid
On Error Resume Next ' Enable error handling
clsid = WshShell.RegRead(regPath)
If Err.Number <> 0 Then
    WScript.Echo "Error reading CLSID for ProgID '" & progID & "'. Error: " & Err.Description
    Err.Clear
Else
    WScript.Echo "CLSID for '" & progID & "' is: " & clsid' #TODO:
End If

' Clean up
Set WshShell = Nothing

' |                              |
' ||||||||||||||||||||||||||||||||
' | Añade Componente a COMapp    |
' ||||||||||||||||||||||||||||||||
' |                              |


' Create an instance of the COM+ Admin catalog
Set catalog = CreateObject("COMAdmin.COMAdminCatalog")

' Get the collection of COM+ Applications
Set apps = catalog.GetCollection("Applications")
apps.Populate

' The name of the COM+ Application you want to add the component to
Dim applicationName
applicationName = "masivo"

' Iterate through the applications to find the target application
For Each app In apps
    If app.Name = applicationName Then
        ' Get the components of the application
        Set components = apps.GetCollection("Components", app.Key)
        components.Populate

        ' Create a new component
        Dim newComponent
        Set newComponent = components.Add
        ' N TODO: Nombre
        newComponent.Name = "comPOLCOMPAG_TX.impCompag" 
        ' newComponent.ComponentCLSID = shell.RegRead("HKEY_CLASSES_ROOT\CLSID\{Your-Component-CLSID}\InprocServer32\")
        newComponent.ComponentCLSID = shell.RegRead("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\InprocServer32\")

        
        ' Save the new component
        components.SaveChanges

        WScript.Echo "Component has been added to '" & applicationName & "'."
        Exit For
    End If
Next

' Clean up
Set newComponent = Nothing
Set components = Nothing
Set app = Nothing
Set apps = Nothing
Set catalog = Nothing
Set shell = Nothing
