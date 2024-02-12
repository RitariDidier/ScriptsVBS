Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
' Set outFile = fso.CreateTextFile("prueba.txt", True)

Dim outputFilePath
outputFilePath = "D:\PipelineCompaginador\pruebaVBS.txt"
Set outputFile = fso.CreateTextFile(outputFilePath, True)

outFile.WriteLine("prueba de ejecucion script vbs Correcta")
WScript.Echo "prueba de ejecucion script vbs Correcta"
outFile.Close
