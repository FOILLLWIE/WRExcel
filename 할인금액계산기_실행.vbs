Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
BasePath = FSO.GetParentFolderName(WScript.ScriptFullName)
SystemPath = BasePath & "\system"

WshShell.Run "powershell -NoProfile -WindowStyle Hidden -Command ""$serverPids = netstat -ano | Select-String ':8000' | ForEach-Object { ($_ -split '\s+')[-1] } | Sort-Object -Unique; foreach ($serverPid in $serverPids) { if ($serverPid -match '^\d+$' -and $serverPid -ne '0') { Stop-Process -Id $serverPid -Force -ErrorAction SilentlyContinue } }""", 0, True
WshShell.CurrentDirectory = SystemPath
WshShell.Run """C:\Users\kwuri\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"" """ & SystemPath & "\server.py""", 0, False
WScript.Sleep 800
WshShell.Run """http://127.0.0.1:8000/index.html""", 1, False
