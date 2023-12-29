Dim fso
Dim sourceFilePath
Dim destinationFilePath

Set fso = CreateObject("Scripting.FileSystemObject")

' sourceFilePath = "C:\\Users\\Administrator\\Desktop\\sshfolder\\ComPolCompag_tx.dll"
' destinationFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"
sourceFilePath = "C:\\Users\\Administrator\\Desktop\\sshfolder\\doc.txt"
destinationFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\doc.txt"

If fso.FileExists(sourceFilePath) Then
    fso.CopyFile sourceFilePath, destinationFilePath
    ' WScript.Echo "File copied successfully."
    WScript.Quit(0)
Else
    WScript.Echo "Source file does not exist."
End If

Set fso = Nothing

