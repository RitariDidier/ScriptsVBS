Dim fso
Dim sourceFilePath
Dim destinationFilePath

Dim dateParts
dateParts = Split(CStr(Date()), "/")

' Assuming the format is DD/MM/YYYY, you would then have:
Dim day, month, year
day = dateParts(0)
month = dateParts(1)
year = dateParts(2)

WScript.Echo "Day: " & day
WScript.Echo "Month: " & month
WScript.Echo "Year: " & year

formattedDate = day & month & year

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Define the source file path and the destination file path
sourceFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"
' ' destinationFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\Backup\\ComPolCompag_tx_"&formattedDate".dll"
' destinationFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\Backup\\ComPolCompag_tx_" & formattedDate & ".dll"
destinationFilePath = "C:\Users\Administrator\Desktop\masivo\Backup\ComPolCompag_tx_" & formattedDate & ".dll"


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

