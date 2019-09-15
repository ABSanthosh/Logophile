Const FONTS = &H14& 
 
Set objShell = CreateObject("Shell.Application") 
Set objFolder = objShell.Namespace(FONTS)

Dim Arg, var1
Set Arg = WScript.Arguments
var1 = Arg(0)


objFolder.CopyHere var1
