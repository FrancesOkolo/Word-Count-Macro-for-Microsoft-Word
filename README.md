# Word Count Macro for Microsoft Word

## Description
This repository contains a VBA (Visual Basic for Applications) macro for Microsoft Word. The macro is designed to:
- Count the number of words in spans under Heading 1 and Heading 2 and Heading 3.
- Calculate the total number of words in a document section from "ABSTRACT" to "REFERENCES".

## Usage
To use the macro, follow these steps:
1. **Open Microsoft Word** and the document you want to analyze.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor, go to `Insert` > `Module` to create a new module.
4. Copy the code from the `WordCountMacro.vbs` file in this repository.
5. Paste the code into the new module in the VBA editor.
6. Close the VBA editor and return to your Word document.
7. Run the macro by pressing `Alt + F8`, selecting `ExportWordCountToCSV`, and clicking `Run`.
8. The macro will generate a CSV file on your desktop containing the word counts.

## Compatibility
This macro is compatible with Microsoft Word. It has been tested on Word versions 2016 and newer. Please ensure that macros are enabled in your Word settings before running this script.
