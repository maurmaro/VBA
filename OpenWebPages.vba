' Function to open a web page in the default browser
Sub OpenWebPages(URL As String)
    Dim ShellCommand As String
    ShellCommand = "explorer.exe " & URL
    Call Shell(ShellCommand, vbNormalFocus)
End Sub
