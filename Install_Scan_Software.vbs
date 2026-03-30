' VBScript Installer for Scan Software
' Run this file as Administrator to install

Option Explicit
Dim objShell, objFSO, strInstallPath, strExePath, strShortcutPath
Dim strSourceExe, objShortcut, objLink, strScriptDir

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Define installation paths
strInstallPath = objShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Scan Software"
strExePath = strInstallPath & "\Scan Software.exe"
strShortcutPath = objShell.ExpandEnvironmentStrings("%ProgramData%") & _
    "\Microsoft\Windows\Start Menu\Programs\Scan Software.lnk"

' Source exe location (where this script is run from)
strSourceExe = strScriptDir & "\dist\Scan Software.exe"

' Check if running as administrator
On Error Resume Next
objFSO.CreateFolder strInstallPath
If Err.Number <> 0 Then
    MsgBox "This installer must be run as Administrator!" & vbCrLf & _
        "Please right-click this file and select 'Run as administrator'", _
        vbCritical, "Admin Rights Required"
    WScript.Quit 1
End If
On Error GoTo 0

' Check if source exe exists
If Not objFSO.FileExists(strSourceExe) Then
    MsgBox "Error: Could not find Scan Software.exe at:" & vbCrLf & strSourceExe, _
        vbCritical, "Installation Failed"
    WScript.Quit 1
End If

' Create installation directory
On Error Resume Next
objFSO.CreateFolder strInstallPath
On Error GoTo 0

' Copy exe to installation directory
On Error Resume Next
objFSO.CopyFile strSourceExe, strExePath, True
If Err.Number <> 0 Then
    MsgBox "Error: Could not copy file to installation directory." & vbCrLf & _
        "Make sure you have administrator privileges.", _
        vbCritical, "Installation Failed"
    WScript.Quit 1
End If
On Error GoTo 0

' Create Start Menu shortcut
On Error Resume Next
Set objLink = objShell.CreateShortcut(strShortcutPath)
objLink.TargetPath = strExePath
objLink.WorkingDirectory = strInstallPath
objLink.Description = "Scan Software - Compliance Scanner"
objLink.IconLocation = strExePath
objLink.Save
On Error GoTo 0

' Show success message
MsgBox "Installation Complete!" & vbCrLf & vbCrLf & _
    "Scan Software has been installed to:" & vbCrLf & strInstallPath & vbCrLf & vbCrLf & _
    "You can now launch it from:" & vbCrLf & _
    "- Start Menu (search for 'Scan Software')" & vbCrLf & _
    "- File Explorer at: " & strInstallPath, _
    vbInformation, "Installation Successful"
