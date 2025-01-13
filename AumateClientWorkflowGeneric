Public Sub AutomateClientWorkflowGeneric()

    '*******************************DEFINE VARIABLES*******************************
    Dim BasePath As String
    Dim Client As String
    Dim Mail As String
    Dim Today As String
    Dim SelectedMail As Outlook.MailItem
    Dim Attachment As Outlook.Attachment
    Dim AttachmentsPath As String
    Dim AttachCount As Integer
    
    ' Declare folder paths
    Dim ClientFolder As String
    Dim MailFolder As String
    
    ' Prompt for Base Path
    BasePath = InputBox("Enter the base path for creating client folders:", "Base Path", "C:\Default\Path\")
    If BasePath = "" Then
        MsgBox "Base path is required. Process aborted.", vbExclamation
        Exit Sub
    End If
    
    ' Ensure path ends with a backslash
    If Right(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
    
    ' Initialize mail subfolder name
    Mail = "\MAIL"
    
    ' Date initialization
    Today = Format(Date, "yyyyMMdd") ' Format: YYYYMMDD
    
    '*******************************************************************************

    '*******************************INPUT*******************************************
    ' Prompt the user for a client name
    Client = InputBox("Enter client name:", "Client Name Folder")
    
    If Client = "" Then
        MsgBox "Client name is required. Process aborted.", vbExclamation
        Exit Sub
    End If
    
    ' Define folder paths
    ClientFolder = BasePath & Client
    MailFolder = ClientFolder & Mail
    
    ' Create client folder and subfolders
    CreateFolder ClientFolder
    CreateFolder MailFolder
    
    '*******************************************************************************

    '*******************************SAVE EMAIL ATTACHMENTS***************************
    ' Check if an email is selected
    On Error Resume Next
    Set SelectedMail = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    
    If SelectedMail Is Nothing Then
        MsgBox "Please select an email to process.", vbExclamation
        Exit Sub
    End If

    ' Save attachments from the selected email
    AttachCount = 0
    For Each Attachment In SelectedMail.Attachments
        AttachmentsPath = ClientFolder & "\" & Today & "_" & Attachment.FileName
        Attachment.SaveAsFile AttachmentsPath
        AttachCount = AttachCount + 1
    Next Attachment
    
    If AttachCount > 0 Then
        MsgBox AttachCount & " attachments saved to: " & ClientFolder, vbInformation
    Else
        MsgBox "No attachments found in the selected email.", vbInformation
    End If
    '*******************************************************************************

    '*******************************OPEN WEB PAGES**********************************
    Dim Links(1 To 2) As String
    Dim i As Integer
    
    ' Prompt user for links
    Links(1) = InputBox("Enter the first link to open:", "Web Link 1", "https://default.link1.com")
    Links(2) = InputBox("Enter the second link to open:", "Web Link 2", "https://default.link2.com")
    
    For i = LBound(Links) To UBound(Links)
        If Links(i) <> "" Then OpenWebPages Links(i)
    Next i
    '*******************************************************************************

    ' Notify user of completion
    MsgBox "Client workflow completed for: " & Client, vbInformation

End Sub

'*******************************HELPER FUNCTIONS*********************************



