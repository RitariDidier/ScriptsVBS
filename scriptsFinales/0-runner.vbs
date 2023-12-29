Set objShell = CreateObject("WScript.Shell")
strScriptPath = "C:\\Users\\Administrator\\Desktop\\scripts\\scriptsVBS\\scriptsFinales\\" 
arrScripts = Array("1-backup.vbs", "2-stopApp.vbs", "3-deleteComponentApp.vbs", "4-Copy.vbs", "5-createComponent.vbs", "6-startApp.vbs")

Dim scriptExitCode
For Each strScript in arrScripts
    scriptExitCode = objShell.Run("cscript """ & strScriptPath & strScript & """", 1, True)
    
    If scriptExitCode = 0 Then
        WScript.Echo strScript & " completed successfully."
    Else
        WScript.Echo strScript & " failed with exit code: " & scriptExitCode
        ' Optional: Exit the loop if a script fails
        ' Exit For
    End If
Next

' Clean up
Set objShell = Nothing
