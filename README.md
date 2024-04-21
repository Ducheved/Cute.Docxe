# DocConverter

## Overview
DocConverter is a multi-threaded document converter that simultaneously converts old .doc files into .docx and .xml formats. It uses the Microsoft Office Interop Word library to open and convert the documents. This project is designed to be efficient and fast, utilizing multiple threads to process multiple files at the same time.

## Requirements
- Microsoft Office must be installed on your machine.
- .NET Framework 4.7.2 or higher is required.

## Usage
1. 1. Clone the repository to your local machine.
2. Open the solution in Visual Studio.
3. Build the solution.
4. Run the program. You will be prompted to enter the path to the folder containing the .doc files you wish to convert.
5. Enter the path to the folder where you want the converted .docx and .xml files to be saved. The program will then convert all the .doc files in the input folder into .docx and .xml formats and save them in the specified output folder. The output files will be organized into separate subfolders for .docx and .xml files.
7. If any files fail to convert, a report will be generated in the output folder detailing which files failed and why.

## Author: Ducheved

## Version: 1.0