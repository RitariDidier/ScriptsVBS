Dim fso, filePath
Dim shell, scriptToRun, arguments

Set fso = CreateObject("Scripting.FileSystemObject")

' ARCHIVO A validar existencia
' filePath = "D:\comPOLCOMPAG_TX"
filePath = "D:\PipelineCompaginador\compilado\comPOLCOMPAG_TX"

' If fso.FileExists(filePath) Then
'     ' Delete the file
'     fso.DeleteFile filePath, True
'     WScript.Echo "File deleted: " & filePath
' Else
'     WScript.Echo "File not found: " & filePath
' End If

' SCRIPT DE COMPILACION DE VBS
scriptToRun = "C:\Program Files\Microsoft Visual Studio\vb6.exe" 

' arguments
' arguments = """bamboobuildworkingdirectory/src/comPOLCOMPAG_TX.vbp"" /m /out ""d:\comPOLCOMPAG_TX\logs\comPOLCOMPAG_TX.log"" /outdir ""D:\comPOLCOMPAG_TX"""
arguments = """D:/PipelineCompaginador/clone/repo/src/comPOLCOMPAG_TX.vbp"" /m /out ""D:\PipelineCompaginador\logs\comPOLCOMPAG_TX.log"" /outdir ""D:\PipelineCompaginador\compilado\comPOLCOMPAG_TX"""
                        ' rutaClone + src + comPOLCOMPAG_TX.vbp                           RutaLOG                                         RutaDLLfinal
Set shell = CreateObject("WScript.Shell")

' corre comando 
shell.Run """" & scriptToRun & """ " & arguments, 1, True

' Release the objects
Set fso = Nothing
Set shell = Nothing


' C:\Program Files\Microsoft Visual Studio\vb6.exe
' ${bamboo.build.working.directory}/src/${bamboo.projectName}.vbp /m /out d:\comPOLCOMPAG_TX\logs\${bamboo.projectName}.log /outdir D://comPOLCOMPAG_TX

' Rutas que necesito

' vbsScript    =  D:\PipelineCompaginador
' Clone        =  D:\PipelineCompaginador\clone
' LOG          =  D:\PipelineCompaginador\logs
' compilado    =  D:\PipelineCompaginador\compilado

' Rutas que existen
' Clone        =   D:/bamboo-agent01-home/xml-data/build-dir/BI-COM-JOB1/src/comPOLCOMPAG_TX.vbp
' LOG          =   d:\comPOLCOMPAG_TX\logs\comPOLCOMPAG_TX.log
' Compilacion  =   D:\comPOLCOMPAG_TX
