' Runs the command line given as its arguments with no window, waits for it, and exits with
' its exit code. The OneDrive Doctor tasks start through this (wscript.exe has no console):
' a console program such as pwsh, started straight from a task with an interactive logon
' type, opens a window on the user's screen every run.
' Ceiling: VBScript is an optional Windows feature, on by default in Windows 11 today; if it
' is ever turned off, these tasks stop running and the installer needs another launcher.
Dim args, i, cmd
Set args = WScript.Arguments
For i = 0 To args.Count - 1
  cmd = cmd & " """ & args(i) & """"
Next
WScript.Quit CreateObject("WScript.Shell").Run(Trim(cmd), 0, True)
