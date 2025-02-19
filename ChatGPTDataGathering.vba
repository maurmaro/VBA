Option Explicit

Sub FetchAndFillCompanyData()
    Dim http As Object
    Dim JSON As Object
    Dim API_URL As String
    Dim API_KEY As String
    Dim VAT_NUMBER As String
    Dim ws As Worksheet
    Dim responseText As String
    
    ' Define your OpenAI API key (Replace with your actual API key)
    API_KEY = "Replace with your actual API key"
    
    ' Set the VAT number you want to query
    VAT_NUMBER = "PUT A VAT NUMBER HERE" ' You can change this dynamically
    
    ' OpenAI API URL for ChatGPT (Example prompt; adjust as needed)
    API_URL = "https://api.openai.com/v1/chat/completions"
    
    ' Create HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Prepare JSON payload
    Dim requestBody As String
    requestBody = "{""model"":""gpt-4"", ""messages"":[{""role"":""system"",""content"":""Provide company details for VAT number " & VAT_NUMBER & """}],""max_tokens"":500}"

    ' Make API request
    With http
        .Open "POST", API_URL, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & API_KEY
        .Send requestBody
        responseText = .responseText
    End With
    
    ' Parse JSON response
    Set JSON = JsonConverter.ParseJson(responseText) ' Requires JSON parser

    ' Define worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Modify if needed
    
    ' Clear previous data
    ws.Cells.Clear

    ' Write data (Modify according to actual API response)
    ws.Cells(1, 1).value = "Company Name"
    ws.Cells(1, 2).value = JSON("choices")(1)("message")("content") ' Extract relevant data
    ws.Cells(2, 1).value = "VAT Number"
    ws.Cells(2, 2).value = VAT_NUMBER
    ' Add more fields if needed...

    ' Notify user
    MsgBox "Company details have been successfully retrieved!", vbInformation, "Success"

End Sub

