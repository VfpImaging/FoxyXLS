# FoxyXLS
Visual FoxPro class to geneate pure XLS files using the BIFF3 file format.
FoxyXLS is a free class that allows Visual FoxPro 9 users to create pure XLS files without the need of having MS-OFFICE or OPENOFFICE installed. It does not use any kind of automation. 

FoxyXLS creates XLS files in the BIFF3 file format directly.

This first version is based on Serhiy Perevoznik class, originally created in Visual C#, adapted to VFP. Some few methods and properties have been included to allow setting the default Column Widths and default Row Heights.

For the future, I hope to work with the Excel 97 file format, that provides more options, and allowing adding formulas.


With the current ALPHA version, with FoxyXLS you can do the following:
Create simple worksheets
Create cells with strings and values
Use several formatting available for dates and numbers
Forecolors and backcolors available
Alignments
All fonts and styles available
Set the default row height and column widths
Set column widths for specific columns

Limitations - basically, the limitations of the MS-Excel 3.0 format:
Only one worksheet page
Formulas are not available yet
Only 16 basic Excel colors available


Additional information on BIFF format and creating XLS files:
Serhiy Perevoznik C# class to create XLS http://delphi32.blogspot.com.br/2011/06/generate-excel-files-without-using.html
Document 'Open Office Excel file format' http://www.openoffice.org/sc/excelfileformat.pdf
Document 'MS Excel 97-2007 binary file format specification' http://download.microsoft.com/download/0/B/E/0BE8BDD7-E5E8-422A-ABFD-4342ED7AD886/Excel97-2007BinaryFileFormat(xls)Specification.pdf


Special thanks to:
Serhiy Perevoznik for his great class http://delphi32.blogspot.com.br/2011/06/generate-excel-files-without-using.html


Excel is a registered trademark of Microsoft Corporation.
