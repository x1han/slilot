' Slilot - hidden autostart for the local service.
' Create a SHORTCUT to this file (keep the file itself inside the repo) and put
' the shortcut into the Startup folder (Win+R -> shell:startup). On login the
' local service starts in a hidden window. The working directory is derived
' from this file's own location, so the repo can live at any path.
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
sh.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
sh.Run "node server.js", 0, False
