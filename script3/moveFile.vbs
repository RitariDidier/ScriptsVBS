Dim fso
Dim sourceFilePath
Dim destinationFilePath

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Define the source file path and the destination file path
sourceFilePath = "C:\\Users\\Administrator\\Desktop\\sshFolder\\ComPolCompag_tx.dll"
destinationFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"

' Check if the source file exists
If fso.FileExists(sourceFilePath) Then
    ' Check if the destination file already exists and delete it if necessary
    ' If fso.FileExists(destinationFilePath) Then
    '     ' Optionally, confirm before overwriting the existing file
    '     fso.DeleteFile destinationFilePath
    ' End If
    ' Copy the file
    fso.CopyFile sourceFilePath, destinationFilePath
    WScript.Echo "File copied successfully."
Else
    WScript.Echo "Source file does not exist."
End If

' Clean up
Set fso = Nothing
