Set WshShell = CreateObject("WScript.Shell")
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""HyperX-Flight"" 1", 0, True) ' Default Device
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setdefaultsounddevice ""HyperX-Flight"" 2", 0, True) ' Default Communications Device
Set WshShell = Nothing