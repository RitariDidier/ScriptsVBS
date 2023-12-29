Dim fso
Dim sourceFilePath
Dim destinationFilePath

Dim dateParts
dateParts = Split(CStr(Date()), "/")

Dim day, month, year
month = dateParts(0)
day = dateParts(1)
year = dateParts(2)

WScript.Echo "Day: " & day
WScript.Echo "Month: " & month
WScript.Echo "Year: " & year

formattedDate = day & month & year

Set fso = CreateObject("Scripting.FileSystemObject")

sourceFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"
' ' destinationFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\Backup\\ComPolCompag_tx_"&formattedDate".dll"
' destinationFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\Backup\\ComPolCompag_tx_" & formattedDate & ".dll"
' destinationFilePath = "C:\Users\Administrator\Desktop\masivo\Backup\ComPolCompag_tx_" & formattedDate & ".dll"


If fso.FileExists(sourceFilePath) Then
    For i = 1 To 15
        WScript.Echo i
        destinationFilePath = "C:\Users\Administrator\Desktop\masivo\Backup\ComPolCompag_tx_" & formattedDate & "-" & i & ".dll"
        If fso.FileExists(destinationFilePath) Then
            ' fso.CopyFile sourceFilePath, destinationFilePath
            WScript.Echo("Existe")
        else 
            ' fso.CopyFile sourceFilePath, destinationFilePath
            WScript.Echo("No Existe, Creando Archivo: " & destinationFilePath)
            Exit For
        End If
    Next
Else
    WScript.Echo "Source file does not exist."
End If

' If fso.FileExists(sourceFilePath) Then
'     If fso.FileExists(destinationFilePath) Then
'         dim version
'         version = "1"
'         destinationFilePath = "C:\Users\Administrator\Desktop\masivo\Backup\ComPolCompag_tx_" & formattedDate & "-" & version & ".dll"
'         fso.CopyFile sourceFilePath, destinationFilePath
'     else 
'         fso.CopyFile sourceFilePath, destinationFilePath
'         WScript.Echo "File copied successfully."
'     End If

' Else
'     WScript.Echo "Source file does not exist."
' End If
' If fso.FileExists(sourceFilePath) Then
'     If fso.FileExists(destinationFilePath) Then
'         dim version
'         version = "1"
'         destinationFilePath = "C:\Users\Administrator\Desktop\masivo\Backup\ComPolCompag_tx_" & formattedDate & "-" & version & ".dll"
'         fso.CopyFile sourceFilePath, destinationFilePath
'     else 
'         fso.CopyFile sourceFilePath, destinationFilePath
'         WScript.Echo "File copied successfully."
'     End If

' Else
'     WScript.Echo "Source file does not exist."
' End If

Set fso = Nothing

