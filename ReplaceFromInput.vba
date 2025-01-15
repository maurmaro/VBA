Sub ReplacePlaceholdersInDocument()
    
    
    Dim FieldPrompts As Variant ' Array to store prompts
    Dim FieldPlaceholders As Variant ' Array to store placeholders
    Dim FieldValues() As String ' Array to store user inputs
    Dim i As Integer ' Loop counter
    
        ' Define prompts in an array for each input field
    ' Define the placeholders corresponding to the input fields
    FieldPlaceholders = Array( _
        "[PolicyNumber]", _
        "[InceptionDate]", _
        "[ClientName]", _
        "[ClientAddress]", _
        "[ClientCAP_City]", _
        "[ClientCountry]", _
        "[ClientPIVA]", _
        "[EndData]", _
        "[RenewalDate]", _
        "[Premium]", _
        "[Limit]", _
        "[BrokerName]", _
        "[BrokerAddress]", _
        "[BrokerCommissions]", _
        "[Deductible11A]", _
        "[Deductible11B]", _
        "[Deductible11C]", _
        "[Deductible11D]", _
        "[Deductible11E]", _
        "[Deductible11F]", _
        "[Deductible12A]", _
        "[Deductible12B]", _
        "[Deductible12C]" _
    )
     ' Define prompts for each input field
    FieldPrompts = Array( _
        "Enter the Policy Number:", _
        "Enter the Inception Date:", _
        "Enter the Client Name:", _
        "Enter the Client Address:", _
        "Enter the Client CAP and City:", _
        "Enter the Client Country:", _
        "Enter the Client PIVA:", _
        "Enter the End Date:", _
        "Enter the Renewal Date:", _
        "Enter the Premium Amount:", _
        "Enter the Policy Limit:", _
        "Enter the Broker Name:", _
        "Enter the Broker Address:", _
        "Enter the Broker Commissions:", _
        "Enter Deductible 11A:", _
        "Enter Deductible 11B:", _
        "Enter Deductible 11C:", _
        "Enter Deductible 11D:", _
        "Enter Deductible 11E:", _
        "Enter Deductible 11F:", _
        "Enter Deductible 12A:", _
        "Enter Deductible 12B:", _
        "Enter Deductible 12C:" _
    )
    
        ' Resize FieldValues array to hold the user inputs
    ReDim FieldValues(LBound(FieldPlaceholders) To UBound(FieldPlaceholders))
    
    ' Loop through prompts and get user input
    For i = LBound(FieldPrompts) To UBound(FieldPrompts)
        FieldValues(i) = InputBox(FieldPrompts(i), "Cyber Insurance Input")
    Next i
    
    ' Loop through the document and replace placeholders with user inputs
    For i = LBound(FieldPlaceholders) To UBound(FieldPlaceholders)
        
        
        ReplaceTextInDocument CStr(FieldPlaceholders(i)), CStr(FieldValues(i))
        
        'MsgBox FieldPlaceholders(i)
        'MsgBox FieldValues(i)
        'ReplaceTextInDocument FieldPlaceholders(i), FieldValues(i)
    Next i
    
    ' Notify the user
    MsgBox "Placeholders replaced successfully!", vbInformation
End Sub

' Function to replace text in the document
Public Sub ReplaceTextInDocument(Placeholder As String, Replacement As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content ' Define the range as the entire document
    
    With rng.Find
        .Text = Placeholder ' Search for the placeholder
        .Replacement.Text = Replacement ' Replace it with user input
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll ' Replace all occurrences
    End With
End Sub
