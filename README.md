# VBA
VBA Code for Automatization
**AutomateClientWorkflowGeneric**
          
This VBA macro automates a series of tasks for handling client-related workflows, such as creating folders, saving email                  attachments, and opening web pages. It is designed to be used in Microsoft Outlook or any VBA environment with email and file             handling capabilities. The macro guides the user through the process with a series of input prompts.

**Open Web Pages**: The user is prompted to enter two URLs (web links). The macro opens both URLs in the default web browser.

**Completion Message**: After all tasks are completed, a message box appears to notify the user that the client workflow has been       successfully completed.**

**Opening Web Pages**: The macro allows the user to open two specified URLs in the default web browser using the OpenWebPages helper     function.

**CreateFolder(FolderPath As String)**: This function checks if a folder exists at the specified FolderPath. If the folder doesn't exist, it creates the folder using MkDir.

**OpenWebPages(URL As String)**: This function opens the provided URL in the system's default web browser using explorer.exe. The Shell function is used to execute this command.

**CopyExcelValueToWordTemplate.vbs**
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
