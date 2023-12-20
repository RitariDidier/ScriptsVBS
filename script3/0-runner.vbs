Set objShell = CreateObject("WScript.Shell")

strScriptPath = "C:\\Users\\Administrator\\Desktop\\scriptsVBS\\scripts3\\" 

arrScripts = Array("1-ListCOMApp.vbs", "2-findComponents.vbs", "3-deleteComponentApp.vbs", "4-createComponent.vbs", "5-startApp.vbs")

For Each strScript in arrScripts
    objShell.Run "cscript """ & strScriptPath & strScript & """", 1, True
Next

' Clean up
Set objShell = Nothing
