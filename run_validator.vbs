' Runs the Streamlit app silently (no black terminal window flashing)
Dim shell
Set shell = CreateObject("WScript.Shell")

' Get folder where this .vbs file is located
Dim scriptDir
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Run batch file hidden
shell.Run "cmd /c cd /d """ & scriptDir & """ && streamlit run app.py --server.headless false --browser.gatherUsageStats false", 0, False
