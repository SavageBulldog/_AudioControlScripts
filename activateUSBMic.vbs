Set WshShell = CreateObject("WScript.Shell")
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""USB Microphone | fifine"" 1", 0, True)
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""USB Microphone | fifine"" 2", 0, True)
Set WshShell = Nothing