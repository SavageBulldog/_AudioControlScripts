Set WshShell = CreateObject("WScript.Shell")
cmds=WshShell.RUN("F:\_RIGTOOLS\nircmd-x64\nircmd.exe setsysvolume 23000", 0, True)
Set WshShell = Nothing