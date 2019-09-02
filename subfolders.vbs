Sub subCreateFolders(strPath)
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    strNewFolder = ""
    Do Until strPath = strNewFolder
        strNewFolder = Left(strPath, InStr(Len(strNewFolder) + 1, strPath, "\"))
        If objFSO.FolderExists(strNewFolder) = False Then
            objFSO.CreateFolder(strNewFolder)
        End If
    Loop
End Sub

Call subCreateFolders("D:\599CD\Downloads\vbs\")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("\\server1\sharedfolders")
Set objSubFolders = objFolder.SubFolders
 
For each objSubFolder In objSubFolders
    wscript.echo objSubFolder
Next

cscript.exe filename.vbs | CLIP