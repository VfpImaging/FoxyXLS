#DEFINE XLSALIGN_DEFAULT         0
#DEFINE XLSALIGN_LEFT            1
#DEFINE XLSALIGN_CENTER          2
#DEFINE XLSALIGN_RIGHT           3

#DEFINE XLSCOLOR_White           1
#DEFINE XLSCOLOR_Red             2
#DEFINE XLSCOLOR_Green           3
#DEFINE XLSCOLOR_Blue            4
#DEFINE XLSCOLOR_Yellow          5
#DEFINE XLSCOLOR_Magenta         6
#DEFINE XLSCOLOR_Cyan            7
#DEFINE XLSCOLOR_DarkRed         8
#DEFINE XLSCOLOR_DarkGreen       9
#DEFINE XLSCOLOR_DarkBlue        10
#DEFINE XLSCOLOR_Olive           11
#DEFINE XLSCOLOR_Purple          12
#DEFINE XLSCOLOR_Teal            13
#DEFINE XLSCOLOR_Silver          14
#DEFINE XLSCOLOR_Gray            15
#DEFINE XLSCOLOR_Black           16
#DEFINE XLSCOLOR_Automatic       17

LOCAL lcFile
lcFile = "c:\Test1.xls"

LOCAL loExcel
SET PROCEDURE TO "FoxyXLS.prg"
loExcel = CREATEOBJECT("FoxyXLS")

loExcel.cAuthor = "VFPIMAGING"
loExcel.nCodePage = 1252

* AddCell(tnRow, tnCol, tuValue, tcFontFull, tcFormat, tnAlign, tnForeColor, tnBackColor)

loExcel.AddCell( 1,1,"Red"      ,"Segoe UI,10,B,Red")
loExcel.AddCell( 2,1,"Green"    ,"Segoe UI,10,B,Green")
loExcel.AddCell( 3,5,"Blue"     ,"Segoe UI,10,B,Blue")

loExcel.nDefaultRowHeight   = 30 && Points
loExcel.nDefaultColumnWidth = 14 && Characters

loExcel.SetColumnWidth(3, 180)
loExcel.SetColumnWidth(1, 180)


loExcel.AddCell( 4,1,"Yellow"   ,"Segoe UI,10,B,Yellow")
loExcel.AddCell( 5,1,"Magenta"  ,"Segoe UI,10,B,Magenta")
loExcel.AddCell( 6,1,"Cyan"     ,"Segoe UI,10,B,Cyan")
loExcel.AddCell( 7,1,"DarkRed"  ,"Segoe UI,10,B,DarkRed")
loExcel.AddCell( 8,1,"DarkGreen","Segoe UI,10,B,DarkGreen")
loExcel.AddCell( 9,1,"DarkBlue" ,"Segoe UI,10,B,DarkBlue")
loExcel.AddCell(10,1,"Olive"    ,"Segoe UI,10,B,Olive")
loExcel.AddCell(11,1,"Purple"   ,"Segoe UI,10,B,Purple")
loExcel.AddCell(12,1,"Teal"     ,"Segoe UI,10,B,Teal")
loExcel.AddCell(13,1,"Silver"   ,"Segoe UI,10,B,Silver")
loExcel.AddCell(14,1,"Gray"     ,"Segoe UI,10,B,Gray")
loExcel.AddCell(15,1,"Black"    ,"Segoe UI,10,B,Black")
loExcel.AddCell(16,1,"Automatic","Segoe UI,10,B,Automatic")

loExcel.AddCell(7,2,26.50    ,"Segoe UI,12,B,DarkRed",,XLSALIGN_CENTER)
loExcel.AddCell(7,3,DATE()   ,"Segoe UI,14,B,Green",,XLSALIGN_LEFT)

loExcel.AddCell(1,4,"Formatted fields:","Segoe Ui,12,B")
loExcel.AddCell(2,4,"Regular"          ,"Segoe Ui,12")
loExcel.AddCell(3,4,"Bold"             ,"Segoe Ui,12,B")
loExcel.AddCell(4,4,"Italic"           ,"Segoe Ui,12,I")
loExcel.AddCell(5,4,"Underlined"       ,"Segoe Ui,12,U")
loExcel.AddCell(6,4,"BoldItalic"       ,"Segoe Ui,12,BI")
loExcel.AddCell(7,4,"Superscript"      ,"Segoe Ui,12,BIS")

loExcel.AddCell(8,3,"Enhanced width to fit the date","Segoe UI,10,B,RED",,XLSALIGN_LEFT)

loExcel.AddCell(20,1,"Date in BRITISH format", "SEGOE UI,12,I")
loExcel.AddCell(20,3,DATE(), "SEGOE UI,12,I","dd/mm/yyyy",XLSALIGN_CENTER)

loExcel.AddCell(21,1,"Date in AMERICAN format", "SEGOE UI,12,I")
loExcel.AddCell(21,3,DATE(), "SEGOE UI,12,I","m/d/yy"    ,XLSALIGN_CENTER)

loExcel.AddCell(23,1,"Values", "SEGOE UI,12,I")
loExcel.AddCell(23,3,1500)

loExcel.AddCell(24,1,"Formatted Values")
loExcel.AddCell(24,3,1500,,"#,##0.00")

loExcel.AddCell(25,1,"Currency formatted Values")
loExcel.AddCell(25,3,1500,,"$#,##0.00")
loExcel.AddCell(26,3,-1500,,"$#,##0.00")

loExcel.AddCell(27,1,"Percentage")
loExcel.AddCell(27,3,0.252,,"0.00%")

loExcel.AddCell( 1,6,"Red"      ,"Segoe UI,10,B,Red",,,,XLSCOLOR_Silver)
loExcel.AddCell( 2,6,"Green"    ,"Segoe UI,10,B,Green")
loExcel.AddCell( 3,6,"Blue"     ,"Segoe UI,10,B,Blue")
loExcel.AddCell( 4,6,"Yellow"   ,"Segoe UI,10,B,Yellow",,,,XLSCOLOR_Silver)
loExcel.AddCell( 5,6,"Magenta"  ,"Segoe UI,10,B,Magenta")
loExcel.AddCell( 6,6,"Cyan"     ,"Segoe UI,10,B,Cyan",,,,XLSCOLOR_Silver)
loExcel.AddCell( 7,6,"DarkRed"  ,"Segoe UI,10,B,DarkRed",,,,XLSCOLOR_Silver)
loExcel.AddCell( 8,6,"DarkGreen","Segoe UI,10,B,DarkGreen")
loExcel.AddCell( 9,6,"DarkBlue" ,"Segoe UI,10,B,DarkBlue")
loExcel.AddCell(10,6,"Olive"    ,"Segoe UI,10,B,Olive",,,,XLSCOLOR_DarkGreen)
loExcel.AddCell(11,6,"Purple"   ,"Segoe UI,10,B,Purple")
loExcel.AddCell(12,6,"White"    ,"Segoe UI,10,B,White",,,,XLSCOLOR_Black)
loExcel.AddCell(13,6,"Silver"   ,"Segoe UI,10,B,Silver")
loExcel.AddCell(14,6,"Gray"     ,"Segoe UI,10,B,Gray")
loExcel.AddCell(15,6,"Black"    ,"Segoe UI,10,B,Black",,,,XLSCOLOR_Green)
loExcel.AddCell(16,6,"Automatic","Segoe UI,10,B,Automatic")


loExcel.WriteFile(lcFile)
RUN /N Explorer.Exe &lcFile.

RETURN