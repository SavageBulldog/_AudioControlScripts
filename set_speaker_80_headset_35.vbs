Set WshShell = CreateObject("WScript.Shell")
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""Speakers"" 1", 0, True) ' Default Device
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setsysvolume 52500", 0, True)
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""HyperX-Flight"" 1", 0, True) ' Default Device -2
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""HyperX-Flight"" 2", 0, True) ' Default Communication Device -2
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setsysvolume 23000", 0, True)
Set WshShell = Nothing
