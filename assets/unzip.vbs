Set ArgObj = WScript.Arguments
Set oShell = WScript.CreateObject ("WScript.Shell")

Dim sCurPath
strFileZIP = "shadow.zip"

If (Wscript.Arguments.Count > 0) Then
  sCurPath = ArgObj(0)
Else
  sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
End if

WScript.Echo(sCurPath)

strZipFile = sCurPath & "\" & strFileZIP

outFolder = sCurPath & "\"

Set objShell = CreateObject( "Shell.Application" )
Set objSource = objShell.NameSpace(strZipFile).Items()
Set objTarget = objShell.NameSpace(outFolder)
intOptions = 256
objTarget.CopyHere objSource, intOptions


