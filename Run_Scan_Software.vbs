' Scan Software Launcher - VBScript Version
' This script copies and runs from user temp folder (no admin rights needed)

Option Explicit
Dim objShell, objFSO, strSourceExe, strTempDir, strTargetExe
Dim strScriptDir

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Define paths - look for exe in SAME DIRECTORY as this script (not a dist subfolder)
strSourceExe = strScriptDir & "\Scan Software.exe"
strTempDir = objShell.ExpandEnvironmentStrings("%TEMP%")
strTargetExe = strTempDir & "\Scan Software.exe"

' Check if source exe exists
If Not objFSO.FileExists(strSourceExe) Then
    MsgBox "Error: Could not find Scan Software.exe" & vbCrLf & vbCrLf & _
        "Expected location: " & strSourceExe & vbCrLf & vbCrLf & _
        "Make sure 'Scan Software.exe' is in the same folder as this script.", _
        vbCritical, "Launch Failed"
    WScript.Quit 1
End If

' Copy exe to temp folder
On Error Resume Next
objFSO.CopyFile strSourceExe, strTargetExe, True
If Err.Number <> 0 Then
    MsgBox "Error: Could not copy to temp folder" & vbCrLf & _
        "This might be a permissions issue.", _
        vbCritical, "Launch Failed"
    WScript.Quit 1
End If
On Error GoTo 0

' Run the exe from temp folder
objShell.Run """" & strTargetExe & """", 1, False

' Inform user
MsgBox "Software launched from temporary location." & vbCrLf & vbCrLf & _
    "The compliance report will open automatically after scanning completes.", _
    vbInformation, "Scan Software Launched"
