' Ask the user to back up the ini files
res=MsgBox("Do you want to back up your ini files?" & Chr(10) & Chr(10) & "If you have done modifications to the .ini files, press 'Yes' and they will be renamed to .ini.bak and will not be deleted.", vbYesNo + vbSystemModal, "Classic Shell Uninstaller")

if res=vbYes then

  Dim objShell
  Set objShell = CreateObject("WScript.Shell") 
  strPath = objShell.RegRead("HKEY_LOCAL_MACHINE\Software\IvoSoft\ClassicShell\Path")

  Dim FSO
  Set FSO = CreateObject("Scripting.FileSystemObject")

  ' Some copy operations may fail because not all ini files exist. So Resume Next.
  On Error Resume Next

  FSO.CopyFile strPath & "\Explorer.ini", strPath & "\Explorer.ini.bak"
  FSO.CopyFile strPath & "\ExplorerL10N.ini", strPath & "\ExplorerL10N.ini.bak"
  FSO.CopyFile strPath & "\StartMenu.ini", strPath & "\StartMenu.ini.bak"
  FSO.CopyFile strPath & "\StartMenuItems.ini", strPath & "\StartMenuItems.ini.bak"
  FSO.CopyFile strPath & "\StartMenuL10N.ini", strPath & "\StartMenuL10N.ini.bak"
end if
