Sub FolderExists(OriginalFile As String, CopyFile As String)
    
    ' Check if the destination file exists
    If Dir(CopyFile) <> "" Then
        ' If the file exists, show a message to inform the user
        MsgBox "The file already exists: " & CopyFile, vbInformation
    Else
        ' If the file doesn't exist, copy the file from the original location to the new location
        On Error GoTo ErrorHandler ' Error handling to manage potential issues with file copy
        FileCopy OriginalFile, CopyFile
        
        ' Inform the user that the file has been successfully copied
        MsgBox "File copied from: " & OriginalFile & vbCrLf & "To: " & CopyFile, vbInformation
        Exit Sub
        
ErrorHandler:
        MsgBox "An error occurred while copying the file. Please check the file paths.", vbCritical
    End If

End Sub
