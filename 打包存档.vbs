Option Explicit

Const SevenZipDir32 = "C:\Program Files\7-Zip\"
Const SevenZipDir64 = "C:\Program Files (x86)\7-Zip\"
Const SevenZipExeName = "7z.exe"
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim sevenZipPath
sevenZipPath = SevenZipDir32 & SevenZipExeName
If Not fso.FileExists(sevenZipPath) Then
    sevenZipPath = SevenZipDir64 & SevenZipExeName
    If Not fso.FileExists(sevenZipPath) Then
        Call MsgBox("7z.exe was not found.", vbExclamation, "Error")
        Call Wscript.Quit(1)
    End If
End If
sevenZipPath = """" & sevenZipPath & """"
Dim shell
Set shell = CreateObject("Wscript.Shell")
Dim status, zipFileName, zipFilePath
Const RootDir = "C:\Users\Administrator\Downloads\"
zipFileName = Year(Now) & "Äê" & Month(Now) & "ÔÂ" & Day(Now) & "ÈÕ" & Hour(Now) & "Ê±" & Minute(Now) & "·Ö" & Second(Now) & "Ãë" & ".7z"
zipFilePath = RootDir & zipFileName
status = shell.Run(sevenZipPath & " a -t7z """ & zipFilePath & """ """ & RootDir & "GrapeCityXa""", , True)
If status <> 0 Then
    Call MsgBox("7z.exe exit code " & status, vbExclamation, "Failed")
Else
    Dim desktop, targetPath
    desktop = shell.SpecialFolders("Desktop")
    targetPath = desktop & "\" & zipFileName
    Call fso.MoveFile(zipFilePath, targetPath)
    Call MsgBox("Backup to " & targetPath, vbInformation, "Done")
End If
