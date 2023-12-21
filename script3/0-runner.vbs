Dim logPath
logPath = "C:\Users\Administrator\Desktop\logs\"

Dim JiraTicket
JiraTicket = "COTVEI-521"

Dim logPathFull
logPathFull = logPath & JiraTicket

WScript.Echo "logPathFull: " & logPathFull

Set objShell = CreateObject("WScript.Shell")
strScriptPath = "C:\\Users\\Administrator\\Desktop\\scripts\\scriptsVBS\\script3\\" 
arrScripts = Array("1-ListCOMApp.vbs", "2-findComponents.vbs", "3-deleteComponentApp.vbs", "4-createComponent.vbs", "5-startApp.vbs")

Dim scriptExitCode
For Each strScript in arrScripts
    scriptExitCode = objShell.Run("cscript """ & strScriptPath & strScript & """", 1, True)
    
    If scriptExitCode = 0 Then
        WScript.Echo strScript & " completed successfully."
    Else
        WScript.Echo strScript & " failed with exit code: " & scriptExitCode
        ' Optional: Exit the loop if a script fails
        Exit For
    End If
Next

' Clean up
Set objShell = Nothing
