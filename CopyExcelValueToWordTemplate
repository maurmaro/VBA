' Module: ModulePolicyDocument

Sub GeneratePolicyDocument()
    ' Define objects and variables
    Dim objWord As Object
    Dim objDoc As Object
    Dim filePath As String
    Dim savePath As String
    Dim currentYear As String
    Dim ClientName As String
    Dim placeholders As Object
    Dim ws As Worksheet

    ' Set worksheet and generic file path
    Set ws = ThisWorkbook.Worksheets("CoverNote") ' Update to match your sheet name
    filePath = InputBox("Enter the template file path:", "File Path", "C:\Templates\FileTemplate.docm")
    
    If filePath = "" Or Dir(filePath) = "" Then
        MsgBox "Invalid file path. Please check and try again.", vbCritical
        Exit Sub
    End If

    ' Get the current year and client name
    currentYear = Year(Date)
    ClientName = ws.Range("ClientName").Value
    savePath = ThisWorkbook.Path & "\" & currentYear & "__RiskPoint_CyberSimplified_" & ClientName & ".docx"

    ' Open Word and the document
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open(filePath)
    objWord.Visible = True

    ' Load placeholders from Excel
    Set placeholders = LoadPlaceholdersFromExcel(ws)

    ' Replace placeholders in Word document
    Call ReplacePlaceholdersInWord(objDoc, placeholders)

    ' Save the document with the new name
    objDoc.SaveAs2 savePath, 16 ' 16 represents wdFormatDocumentDefault

    ' Close Word application
    objDoc.Close
    objWord.Quit

    ' Release objects
    Set objDoc = Nothing
    Set objWord = Nothing

    ' Display success message
    MsgBox "Document created and saved as: " & savePath, vbInformation
End Sub
