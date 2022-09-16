Set WshShell = CreateObject("WScript.Shell")
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""Speakers"" 1", 0, True) ' Default Device
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setsysvolume 23000", 0, True)
Set WshShell = Nothing