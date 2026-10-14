' start_hidden.vbs - Start the server WITHOUT a black window (runs in background).
' Double-click to start. To stop it, double-click stop_server.bat.
' Messages are written to server.log in this folder.
' (Comments are in English on purpose: English Windows may break Korean text here.)

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Full path of run.bat in this file's folder
bat = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "run.bat")

' 0 = hidden window, False = do not wait
shell.Run """" & bat & """ hidden", 0, False
