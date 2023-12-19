' |                              |
' ||||||||||||||||||||||||||||||||
' |         FindProgID           |
' ||||||||||||||||||||||||||||||||
' |                              |
' Set the name of the DLL file
Dim dllName
dllName = "C:\Users\Administrator\Desktop\masivo\ComPolCompag_tx 1.dll"

' Create a shell object
Dim shell
Set shell = CreateObject("WScript.Shell")

' Registry path to start the search
Dim regPath
regPath = "HKEY_CLASSES_ROOT\"

' Command to search the registry
Dim command
command = "reg query " & Chr(34) & regPath & Chr(34) & " /s /f " & Chr(34) & dllName & Chr(34) & " /t REG_SZ"

' Run the command
Dim execObject
Set execObject = shell.Exec(command)

' Check for errors
If execObject.Status = 0 Then
    ' Capture and display the output if no error
    Dim output
    output = execObject.StdOut.ReadAll
    WScript.Echo output
    WScript.Echo "a"
Else
    WScript.Echo "Error in executing registry query."
End If

' Clean up
Set shell = Nothing
