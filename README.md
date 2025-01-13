# VBA
VBA Code for Automatization
AutomateClientWorkflowGeneric: 
          This VBA macro automates a series of tasks for handling client-related workflows, such as creating folders, saving email                  attachments, and opening web pages. It is designed to be used in Microsoft Outlook or any VBA environment with email and file             handling capabilities. The macro guides the user through the process with a series of input prompts.

Main Steps of the Workflow:
          Folder Creation:

                    The user is prompted to enter a base path for creating client folders. It ensures the path ends with a backslash (\).
                    The macro automatically creates two folders: one for the client (Client) and one for the mail attachments (MAIL).
                    
Email Attachment Handling:
                    The user is asked to select an email from their active Outlook window.
                    If the selected email contains attachments, these attachments are saved into the client's folder with a date-stamped                      filename. The number of attachments saved is shown to the user.
Open Web Pages:
                    The user is prompted to enter two URLs (web links). The macro opens both URLs in the default web browser.
Completion Message:
                    After all tasks are completed, a message box appears to notify the user that the client workflow has been                                 successfully completed.
Functionality Breakdown:
                    Folder Creation (CreateFolder helper function):
                    The macro ensures that the folder structure exists. If the folder does not exist, the CreateFolder function is                            invoked to create it.
Saving Email Attachments:
                    The macro checks if there are any attachments in the selected email and saves them to the client's folder with a 
                    date-specific filename.
Opening Web Pages:
                    The macro allows the user to open two specified URLs in the default web browser using the OpenWebPages helper                             function.

Helper Functions:
                    CreateFolder(FolderPath As String):

                    This function checks if a folder exists at the specified FolderPath. If the folder doesn't exist, it creates the                          folder using MkDir.

OpenWebPages(URL As String):

                    This function opens the provided URL in the system's default web browser using explorer.exe. The Shell function is                        used to execute this command.
Usage:
                    Run the AutomateClientWorkflowGeneric subroutine from the VBA editor.
                    Follow the prompts to enter the base path for folders, client name, and URLs.
                    The macro will create the necessary folders, save email attachments, and open the specified web pages.
Notes:
                    The macro relies on the Outlook object model (for selecting emails) and the system's file management system (for                          folder creation and file saving).
                    It also assumes that the user has access to the required file system and that Outlook is properly configured.

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
