Dim fso
Dim shell, scriptToRun, arguments

Set fso = CreateObject("Scripting.FileSystemObject")

scriptToRun = "C:\Program Files\Microsoft Visual Studio\vb6.exe"

arguments = """D:/PipelineCompaginador/clone/repo/src/comPOLCOMPAG_TX.vbp"" /m /out ""D:\PipelineCompaginador\logs\comPOLCOMPAG_TX.log"" /outdir ""D:\PipelineCompaginador\compilado\comPOLCOMPAG_TX"""

Set shell = CreateObject("WScript.Shell")

shell.Run """" & scriptToRun & """ " & arguments, 1, True

Set fso = Nothing
Set shell = Nothing