Dim fso
Dim sourceFilePath
Dim destinationFilePath

Dim dateParts
dateParts = Split(CStr(Date()), "/")

Dim day, month, year
month = dateParts(0)
day = dateParts(1)
year = dateParts(2)

' Check if month is a single digit and prepend "0" if necessary
If Len(month) = 1 Then
    month = "0" & month
End If

formattedDate = day & month & year

Set fso = CreateObject("Scripting.FileSystemObject")

sourceFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\ComPolCompag_tx.dll"
' sourceFilePath = "C:\\Users\\Administrator\\Desktop\\masivo\\doc.txt"


If fso.FileExists(sourceFilePath) Then
    For i = 1 To 15
        destinationFilePath = "C:\Users\Administrator\Desktop\masivo\Backup\ComPolCompag_tx_" & formattedDate & "-" & i & ".dll"
        ' destinationFilePath = "C:\Users\Administrator\Desktop\masivo\Backup\doc" & formattedDate & "-" & i & ".txt"
        If Not fso.FileExists(destinationFilePath) Then
            fso.CopyFile sourceFilePath, destinationFilePath
            ' WScript.Echo("No Existe, Creando Archivo: " & destinationFilePath)
            Exit For    
        End If
    Next
    WScript.Quit(0)
Else
    WScript.Echo "Source file does not exist."
End If

Set fso = Nothing

