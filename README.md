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
          
