# VBA
VBA Code for Automatization

CopyExcelValueToWordTemplate.vbs
          Generic Module:
          
                    ReplacePlaceholdersInWord: Reusable logic for replacing placeholders.
                    LoadPlaceholdersFromExcel: Dynamically populates the dictionary from an Excel sheet.
          
          Specific Procedure:
                    Uses InputBox to allow users to specify the file path dynamically.
                    Handles invalid file paths gracefully with error checking.
                    Saves the document with a structured name based on the current year and client name.
          
          Modular Structure:
          
                    Generic logic is abstracted into ModuleGeneric for reuse across projects.
                    Policy-specific details are confined to ModulePolicyDocument.
          Dynamic Path Selection:
          
                    Prompts the user for the document template path, ensuring flexibility.

FolderExists.vbs:
          Automates manual folder and file operations, saving time and reducing errors.
          Provides reusable code that can be integrated into larger VBA projects.
          Enhances user experience with clear prompts and notifications.
          
          Dependencies:
                    Ensure that all required file paths are accessible and valid.
                    Modify static paths, URLs, or file names to suit the specific environment or project.

OpenWebPages.vbs:
          This module provides a simple function to open a web page in the default web browser using Visual Basic for Applications (VBA).             It utilizes the Shell function to invoke the system's default browser (typically Internet Explorer or whichever browser is set as           default) to open the specified URL. 
          Function: OpenWebPages(URL As String)
          Purpose: The function allows users to open a web page in their default web browser by passing the URL as an argument.

          Parameters:

                    URL (String): The URL of the webpage that you want to open. This should include the full URL (e.g.,           "https://www.example.com").
          How it works:

                    The function constructs a command string that calls explorer.exe, which is the executable for Windows File Explorer                        and can also open URLs in the default browser.
                    The Shell function is then used to execute this command, launching the default web browser with the specified URL.
                    Usage: You can use this function within any VBA project (e.g., in Microsoft Excel, Word, etc.) to quickly open web pages programmatically.
