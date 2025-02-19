' This VBA script for Microsoft Word calculates insurance taxes based on a given [Premium] value and replaces placeholders in the document with the computed amounts. You need to have on the word page the TAGS [Premium], [FirstParty], [ThirdParty], [Taxes], [GrossPremium]   

Sub InsuranceTaxComputation()
    Dim firstParty As Double
    Dim thirdParty As Double
    Dim taxes As Double
    Dim grossPremium As Double
    Dim premium As Double
    Dim totalCoverages As Integer
    Dim firstPartyCoverages As Integer
    Dim thirdPartyCoverages As Integer

    premium = InputBox("Inser the premium gross of brokerage: ")

    ' Calculate the required values
    'Number of firstPartyCoverages
  
    'Put the number below
    totalCoverages = 9
    
    'Put the number below
    firstPartyCoverages = 6

    'Number of thirdPartyCoverages
    'Put the number here
    thirdPartyCoverages = 3
  
    firstParty = (premium / totalCoverages) * firstPartyCoverages * 0.2125
    thirdParty = (premium / totalCoverages) * thirdPartyCoverages * 0.2225
    taxes = firstParty + thirdParty
    
     grossPremium = premium + taxes

    ' Replace placeholders in the document with formatted values (without default currency symbol)
    SostituisciSegnaposto "[Premium]", FormatNumber(premium, 2) & " €"
    SostituisciSegnaposto "[FirstParty]", FormatNumber(firstParty, 2) & " €"
    SostituisciSegnaposto "[ThirdParty]", FormatNumber(thirdParty, 2) & " €"
    SostituisciSegnaposto "[Taxes]", FormatNumber(taxes, 2) & " €"
    SostituisciSegnaposto "[GrossPremium]", FormatNumber(grossPremium, 2) & " €"
End Sub

' Function to find and replace a placeholder in the document with a specific value
Sub SostituisciSegnaposto(segnaposto As String, valore As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = segnaposto
        .Replacement.Text = valore
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Sub

