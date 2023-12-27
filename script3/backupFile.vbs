Dim fso
Dim sourceFilePath
Dim destinationFilePath

Dim currentDate
currentDate = Date()
WScript.Echo "Current Date: " & currentDate ' Add this line for debugging

Dim day, month, year
day = Day(currentDate)
WScript.Echo "Day: " & day ' Add this line for debugging
month = Month(currentDate)
WScript.Echo "Month: " & month ' Add this line for debugging
year = Year(currentDate)
WScript.Echo "Year: " & year ' Add this line for debugging

' Formatting the date as DD/MM/YYYY
Dim formattedDate
formattedDate = Right("0" & day, 2) & "/" & Right("0" & month, 2) & "/" & year
WScript.Echo "formattedDate: " & formattedDate ' Add this line for debugging


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

