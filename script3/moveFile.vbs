Dim fso
Dim sourceFilePath
Dim destinationFilePath

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Define the source file path and the destination file path
' sourceFilePath = "C:\Users\Administrator\Desktop\sshFolder\ComPolCompag_tx.dll"
sourceFilePath = "C:\\Users\\Administrator\\Desktop\\sshFolder\\file.txt"
destinationFilePath = "C:\Users\Administrator\Desktop\masivo"

' Check if the source file exists
If fso.FileExists(sourceFilePath) Then
    ' Move the file
    fso.MoveFile sourceFilePath, destinationFilePath
    WScript.Echo "File moved successfully."
Else
    WScript.Echo "Source file does not exist."
End If

' Clean up
Set fso = Nothing
