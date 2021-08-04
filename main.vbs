Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

m=minute(time)
Select Case m
	Case 01
	objShell.Run "cscript VBScript\CreateFile01.vbs"
	Case 02
	objShell.Run "cscript VBScript\SaveFile02.vbs"
	Case 03
	objShell.Run "cscript VBScript\SaveFile03.vbs"
	Case 04
	objShell.Run "cscript VBScript\SaveFile04.vbs"
	Case 05
	objShell.Run "cscript VBScript\SaveFile05.vbs"
	Case 06
	objShell.Run "cscript VBScript\SaveFile06.vbs"
	Case 07
	objShell.Run "cscript VBScript\SaveFile07.vbs"
	Case 08
	objShell.Run "cscript VBScript\SaveFile08.vbs"
	Case 09
	objShell.Run "cscript VBScript\SaveFile09.vbs"
	Case 10
	objShell.Run "cscript VBScript\SaveFile10.vbs"
	Case else
	objShell.Run "cscript VBScript\SaveFileWrong.vbs"
End Select
