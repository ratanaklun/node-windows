Set Shell = CreateObject("Shell.Application")
Set WShell = WScript.CreateObject("WScript.Shell")
Set ProcEnv = WShell.Environment("PROCESS")

cmd = ProcEnv("CMD")
app = ProcEnv("APP")
args= Right(cmd,(Len(cmd)-Len(app)))

If (WScript.Arguments.Count >= 1) Then
  On Error Resume Next
  WShell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If Err.Number = 0 Then
    Shell.ShellExecute app, args, "", "open", 0
  Else
    Shell.ShellExecute app, args, "", "runas", 0
  End If
Else
  WScript.Quit
End If
