* FOXYXLS Class
* Creates pure XLS files without the need of having MS-OFFICE installed
* Author: VFPIMAGING
* Based on Serhiy Perevoznyk's work published in his blog
* http://delphi32.blogspot.com.br/2011/06/generate-excel-files-without-using.html

#DEFINE DefaultColor         0x7fff

#DEFINE BOFRecord            0x0209
#DEFINE EOFRecord            0x0A

#DEFINE FontRecord           0x0231
#DEFINE FormatRecord         0x001E
#DEFINE LabelRecord          0x0204
#DEFINE WindowProtectRecord  0x0019
#DEFINE XFRecord             0x0243
#DEFINE HeaderRecord         0x0014
#DEFINE FooterRecord         0x0015
#DEFINE ExtendedRecord       0x0243
#DEFINE StyleRecord          0x0293
#DEFINE CodepageRecord       0x0042
#DEFINE NumberRecord         0x0203
#DEFINE ColumnInfoRecord     0x007D

#DEFINE DefColWidth          0x0055
#DEFINE DefRowHeight         0x0225

DEFINE CLASS FoxyXLS AS Custom
	cXLS      = ""
	cFile     = ""
	cDateType = "B"
	cHeader   = ""
	cFooter   = ""
	cAuthor   = ""
	nCodePage           = 1252
	_nCurrFXRecord      = 20
	cDefaultFontName    = "Arial"
	nDefaultFontSize    = 10
	nDefaultColumnWidth = 0
	nDefaultRowHeight   = 0

FUNCTION Init

	This.AddProperty("aFonts[1,1]"    , "")
	This.AddProperty("aCells[1,1]"    , "")
	This.AddProperty("aFmts[1,1]"     , "")
	This.AddProperty("aColWidths[1,1]", "")
	This.AddProperty("aColors[1,1]"   , "")

	This.PrepareFormatArray()
	This.PrepareColorsArray()

	LOCAL lcSetDate, lcDateType
	lcSetDate = UPPER(ALLTRIM(SET("Date")))

	DO CASE
	CASE INLIST(lcSetDate, "BRITISH", "FRENCH", "DMY")
		lcDateType = "B"

	CASE INLIST(lcSetDate, "AMERICAN", "USA", "MDY")
		lcDateType = "A"

	OTHERWISE
		lcDateType = "B"
	ENDCASE
	This.cDateType = lcDateType

	* Prepare the fonts Array
	DIMENSION This.aFonts(4, 2)
	This.aFonts(1,1) = "Arial,10,"
	This.aFonts(2,1) = "Arial,10,B"
	This.aFonts(3,1) = "Arial,10,I"
	This.aFonts(4,1) = "Arial,10,BI"

	This.aFonts(1,2) = 0
	This.aFonts(2,2) = 0
	This.aFonts(3,2) = 0
	This.aFonts(4,2) = 0
ENDFUNC 






FUNCTION WriteFile(tcFile)
	IF EMPTY(tcFile)
		tcFile = This.cFile
	ENDIF 
	This.CreateHeader()

	LOCAL n, lcBiffCol
	lcBiffCol = ""
	

	* Setting the default row heights
	lcBiffCol = lcBiffCol + This.WriteDefaultRowHeight()


	* Setting the default row heights
	*	lcBiffCol = lcBiffCol + ;
		0h0802 + ;  && Index
		0h1000 + ;  && Size
		0h0000 + ;  && Row number
		0h0000 + ;  && Index to column of the first cell which is described by a cell record
		0h0400 + ;  && Index to column of the last cell which is described by a cell record, increased by 1
		BINTOC(40 * 20, "2RS") + ;  && Height of the row, in twips = 1/20 of a point
		0h0000 + ;  && Not used
		0h1022 + ;  && a relative offset to calculate stream position of the first cell record for this row (?4.7.2)
		0h0100f000  && Option flags and default row formatting

	* Setting the default column widths
	lcBiffCol = lcBiffCol + This.WriteDefaultColumnWidth()

	
	* Setting the column forced widths
	IF VARTYPE(This.aColWidths(1,1)) = "N"
		FOR n = 1 TO ALEN(This.aColWidths, 1)
			lcBiffCol = lcBiffCol + This.WriteColumnInfoRecord(This.aColWidths(n, 1), This.aColWidths(n, 2))
		ENDFOR 
	ENDIF 

	* Writing the cells
	FOR n = 1 TO ALEN(This.aCells, 1)
		This.WriteCell(n, This.aCells(n,1), This.aCells(n,2), This.aCells(n,3))	
	ENDFOR 

	This.cXLS = This.cHeader + lcBiffCol + This.cXLS + This.cFooter
	=STRTOFILE(This.cXLS, tcFile)
ENDFUNC 



PROCEDURE GetFontFullName
	LPARAMETERS tnIndex, tcFontFull, tnForeColor, tcColorName, tnXLSColor

	* Check if only the forecolor was passed
	IF EMPTY(tcFontFull) AND tnForeColor > 0
		TRY 
			tcColorName = This.aColors(tnForeColor, 1)
			tnXlsColor  = This.aColors(tnForeColor, 2)
			lcFontFull  = ",,," + tcColorName
		CATCH
		ENDTRY 
		RETURN lcFontFull
	ENDIF 

	
	* Working with the supposition NOT EMPTY(tcFontFull)
		
	* Getting the correct Forecolor value
	* And preparing the internal tables to work with it
	LOCAL lcFontFull, lcColorName, lnColor
	LOCAL lcFont, lnSize, lcStyle, lcColor
	lcFont  = ALLTRIM(GETWORDNUM(tcFontFull, 1, ","))
	IF EMPTY(lcFont)
		lcFont = This.cDefaultFontName
	ENDIF 
	lnSize  = INT(VAL(ALLTRIM(GETWORDNUM(tcFontFull, 2, ","))))
	IF EMPTY(lnSize)
		lnSize = This.nDefaultFontSize
	ENDIF 
	lcStyle = UPPER(ALLTRIM(GETWORDNUM(tcFontFull, 3, ",")))
	lcColor = UPPER(ALLTRIM(GETWORDNUM(tcFontFull, 4, ",")))

	tcFontFull = lcFont + "," + TRANSFORM(lnSize) + "," + lcStyle + ","

	DO CASE
	CASE EMPTY(lcColor) AND tnForeColor = 0 && No color passed
		tcColorName = ""
		tnXlsColor  = 0
		lcFontFull  = tcFontFull

	CASE NOT EMPTY(lcColor)
		TRY 
			lnIndex = CEILING(ASCAN(This.aColors, lcColor, 1, 0, 1,4+2+1)/3)
			IF lnIndex = 0 && Not found, the ignore the color
				tcColorName = ""
				tnXlsColor  = 0
				lcFontFull  = tcFontFull
			ELSE 
				&& Found the color
				tcColorName = lcColor
				tnXlsColor  = This.aColors(lnIndex, 2)
				lcFontFull  = tcFontFull + lcColor
			ENDIF 
		CATCH
			tcColorName = ""
			tnXlsColor  = 0
			lcFontFull  = tcFontFull
		ENDTRY 

	OTHERWISE
		tcColorName = ""
		tnXlsColor  = 0
		lcFontFull  = tcFontFull
	ENDCASE

	RETURN lcFontFull
ENDPROC 



PROCEDURE AddCell(tnRow, tnCol, tuValue, tcFontFull, tcFormat, tnAlign, tnForeColor, tnBackColor)
	tcFontFull  = EVL(tcFontFull, "")
	tnForeColor = EVL(tnForeColor, 0)
	* Prepare the Cells array, having the next record initialized to receive the information
	LOCAL lnTotCells, i
	lnTotCells = ALEN(This.aCells, 1)
	IF lnTotCells = 1 AND VARTYPE(This.aCells(1)) <> "N"
		i = 1
	ELSE 
		i = lnTotCells + 1
	ENDIF 
	DIMENSION This.aCells(i, 10)

	LOCAL lcFontFull, lnXLSColor, lcColorName
	lcFontFull  = ""
	lnXLSColor  = 0
	lcColorName = ""
	IF EMPTY(tcFontFull) AND EMPTY(tnForeColor)
	ELSE
		lcFontFull = This.GetFontFullName(i, tcFontFull, tnForeColor, @lcColorName, @lnXLSColor)
	ENDIF 

	IF EMPTY(lcFontFull)
		lnFontIndex = 0
	ELSE 
		lnFontIndex = CEILING(ASCAN(This.aFonts, lcFontFull,1,0,1,4+2+1)/2)
		IF lnFontIndex = 0
			lnFontTotal = ALEN(This.aFonts,1)
			lnFontIndex = lnFontTotal + 1
			DIMENSION This.aFonts(lnFontIndex,2)
			This.aFonts(lnFontIndex, 1) = lcFontFull
			This.aFonts(lnFontIndex, 2) = lnXLSColor
		ENDIF 
	ENDIF 

	LOCAL lcValueType
	lcValueType = VARTYPE(m.tuValue)
	IF lcValueType = "D" AND EMPTY(tcFormat)
		DO CASE
		CASE This.cDateType = "A"
			tcFormat = "m/d/yy"
		CASE This.cDateType = "B"
			tcFormat = "dd/mm/yyyy"
		OTHERWISE
		ENDCASE
	ENDIF 

	* Getting the correct Format Index
	LOCAL lnFormatIndex, luFmtIndex
	luFmtIndex = tcFormat
	DO CASE
	CASE VARTYPE(luFmtIndex) = "N"
		lnFormatIndex = luFmtIndex
	CASE VARTYPE(luFmtIndex) = "C"
		lnFormatIndex = ASCAN(This.aFmts, ALLTRIM(luFmtIndex),1,0,1,4+2+1) && All records, column 1, Case insensitive
		lnFormatIndex = MAX(lnFormatIndex - 1, 0)
	OTHERWISE
		lnFormatIndex = 0
	ENDCASE

	This.aCells(i, 1) = tnRow
	This.aCells(i, 2) = tnCol
	This.aCells(i, 3) = tuValue
	This.aCells(i, 4) = lnFontIndex
	This.aCells(i, 5) = lnFormatIndex
	This.aCells(i, 6) = tnAlign
	This.aCells(i, 7) = lnXLSColor
	This.aCells(i, 8) = tnBackColor
ENDPROC 


FUNCTION CreateHeader
	This.cFooter = BINTOC(EOFRecord, "4RS")

	LOCAL lcBiffStart, lcBiffAuthor, lcBiffCodePage, lcBiffFontTable, lcBiffHeaderRecord, lcBiffFooterRecord
	LOCAL lcBiffFormatTable, lcBiffWindowProtect, lcBiffXFTable, lcBiffStyle
	lcBiffStart         = This.GetMainHeader()
    lcBiffAuthor        = This.WriteAuthorRecord()
    lcBiffCodePage      = This.WriteCodepageRecord()
    lcBiffFontTable     = This.WriteFontTable()
    lcBiffHeader        = This.WriteHeaderRecord()
    lcBiffFooter        = This.WriteFooterRecord()
    lcBiffFormatTable   = This.WriteFormatTable()
    lcBiffWindowProtect = This.WriteWindowProtectRecord()
    lcBiffXFTable       = This.WriteXFTable()
    lcBiffStyle         = This.WriteStyleTable()

	LOCAL lcFullHeader
	lcFullHeader = lcBiffStart + ;
					lcBiffAuthor + ;
					lcBiffCodePage + ;
					lcBiffFontTable + ;
					lcBiffHeader + ;
					lcBiffFooter + ;
					lcBiffFormatTable + ;
					lcBiffWindowProtect + ;
					lcBiffXFTable + ;
					lcBiffStyle

	This.cHeader = lcFullHeader
ENDFUNC 


FUNCTION GetMainHeader
	LOCAL lcBiffHeader
	lcBiffHeader = 0h090208000000100000000000
	RETURN lcBiffHeader
ENDFUNC 

FUNCTION WriteAuthorRecord
	LOCAL lcBiffAuthor
	lcBiffAuthor = 0h5C00200006 + PADR(This.cAuthor, 31, " ")  && 31 positions
	RETURN lcBiffAuthor
ENDFUNC 

FUNCTION WriteCodepageRecord
	LOCAL lcBiffCodePage
	IF VARTYPE(This.nCodePage) = "C"
		This.nCodePage = VAL(This.nCodePage)
	ENDIF 	
	
	lcBiffCodePage = 0h42000200 + 0h + BINTOC(This.nCodePage,"2RS")
	RETURN lcBiffCodePage
ENDFUNC 


FUNCTION WriteFontTable
	* Write font table
	LOCAL lcBiffFontRecord, n, lcFont
	lcBiffFontRecord = ""
	FOR n = 1 TO ALEN(This.aFonts,1)
		lcFont = This.aFonts(n, 1)
		IF NOT EMPTY(lcFont)
			lcBiffFontRecord = lcBiffFontRecord + This.WriteFontRecord(n, lcFont)
		ENDIF 
	ENDFOR 
	RETURN lcBiffFontRecord
ENDFUNC 


FUNCTION WriteFontRecord(tnIndex, tcFontFull)

	* Getting the correct Forecolor value
	LOCAL lcFont, lnSize, lcStyle, lcArray, lnIntStyle, lcColor, lnColor, lcColorName, lnClrIndex
	lcFont  = ALLTRIM(GETWORDNUM(tcFontFull, 1, ","))
	lnSize  = VAL(ALLTRIM(GETWORDNUM(tcFontFull, 2, ",")))
	lcStyle = UPPER(ALLTRIM(GETWORDNUM(tcFontFull, 3, ",")))
	lnColor = This.aFonts(tnIndex, 2)

	lnIntStyle = IIF("B" $ lcStyle, 1, 0) + ;
		IIF("I" $ lcStyle, 2, 0) + ;
		IIF("U" $ lcStyle, 4, 0) + ;
		IIF("S" $ lcStyle, 8, 0)

	lcArray = BINTOC(FontRecord, "2RS")   + ; && 0h3102 = FontRecord
		BINTOC(LEN(lcFont) + 7,"2RS") + ; && Fontname length + 7
		BINTOC(lnSize * 20, "2RS")    + ; && FontSize
		BINTOC(lnIntStyle, "2RS")     + ; && FontStyle
		BINTOC(lnColor   , "2RS")   + ; && Color
		CHR(LEN(lcFont))              + ;
		lcFont

	RETURN lcArray
ENDFUNC


FUNCTION PrepareColorsArray
	DIMENSION This.aColors(17,3)
	WITH This
		.aColors(1,1) = "WHITE"
		.aColors(1,2) = 0x01    && ForeColor
		.aColors(1,3) = 0xC041  && BackColor

		.aColors(2,1) = "RED"
		.aColors(2,2) = 0x02    && ForeColor
		.aColors(2,3) = 0xC081  && BackColor

		.aColors(3,1) = "GREEN"
		.aColors(3,2) = 0x03    && ForeColor
		.aColors(3,3) = 0xC0C1  && BackColor

		.aColors(4,1) = "BLUE"
		.aColors(4,2) = 0x04    && ForeColor
		.aColors(4,3) = 0xC101  && BackColor

		.aColors(5,1) = "YELLOW"
		.aColors(5,2) = 0x05    && ForeColor
		.aColors(5,3) = 0xC141  && BackColor

		.aColors(6,1) = "MAGENTA"
		.aColors(6,2) = 0x06    && ForeColor
		.aColors(6,3) = 0xC181  && BackColor

		.aColors(7,1) = "CYAN"
		.aColors(7,2) = 0x07    && ForeColor
		.aColors(7,3) = 0xC1C1  && BackColor

		.aColors(8,1) = "DARKRED"
		.aColors(8,2) = 0x10    && ForeColor
		.aColors(8,3) = 0xC401  && BackColor

		.aColors(9,1) = "DARKGREEN"
		.aColors(9,2) = 0x11    && ForeColor
		.aColors(9,3) = 0xC441  && BackColor

		.aColors(10,1) = "DARKBLUE"
		.aColors(10,2) = 0x12    && ForeColor
		.aColors(10,3) = 0xC481  && BackColor

		.aColors(11,1) = "OLIVE"
		.aColors(11,2) = 0x13    && ForeColor
		.aColors(11,3) = 0xC4C1  && BackColor

		.aColors(12,1) = "PURPLE"
		.aColors(12,2) = 0x14    && ForeColor
		.aColors(12,3) = 0xC501  && BackColor

		.aColors(13,1) = "TEAL"
		.aColors(13,2) = 0x15    && ForeColor
		.aColors(13,3) = 0xC541  && BackColor

		.aColors(14,1) = "SILVER"
		.aColors(14,2) = 0x16    && ForeColor
		.aColors(14,3) = 0xC581  && BackColor

		.aColors(15,1) = "GRAY"
		.aColors(15,2) = 0x17    && ForeColor
		.aColors(15,3) = 0xC5C1  && BackColor

		.aColors(16,1) = "BLACK"
		.aColors(16,2) = 0x00    && ForeColor
		.aColors(16,3) = 0xC001  && BackColor

		.aColors(17,1) = "AUTOMATIC"
		.aColors(17,2) = 0x7FFF  && ForeColor
		.aColors(17,3) = 0xCE00  && BackColor
	ENDWITH 
ENDFUNC 



FUNCTION WriteHeaderRecord
	LOCAL lcBiffHeaderRecord
	lcBiffHeaderRecord = BINTOC(HeaderRecord, "4RS")
	RETURN lcBiffHeaderRecord
ENDFUNC 


FUNCTION WriteFooterRecord
	LOCAL lcBiffFooterRecord
	lcBiffFooterRecord = BINTOC(FooterRecord, "4RS")
	RETURN lcBiffFooterRecord
ENDFUNC 


FUNCTION PrepareFormatArray
	* Prepare the Formatting record
	DIMENSION This.aFmts(38)
	WITH This
    .aFmts(1)=("General")
    .aFmts(2)=("0")
    .aFmts(3)=("0.00")
    .aFmts(4)=("#,##0")
    .aFmts(5)=("#,##0.00")
    .aFmts(6)=("($#,##0_);($#,##0)")
    .aFmts(7)=("($#,##0_);[Red]($#,##0)")
    .aFmts(8)=("($#,##0.00_);($#,##0.00)")
    .aFmts(9)=("($#,##0.00_);[Red]($#,##0.00)")
    .aFmts(10)=("0%")
    .aFmts(11)=("0.00%")
    .aFmts(12)=("0.00E+00")
    .aFmts(13)=("# ?/?")
    .aFmts(14)=("# ??/??")
    .aFmts(15)=("m/d/yy")
    .aFmts(16)=("d-mmm-yy")
    .aFmts(17)=("d-mmm")
    .aFmts(18)=("mmm-yy")
    .aFmts(19)=("h:mm AM/PM")
    .aFmts(20)=("h:mm:ss AM/PM")
    .aFmts(21)=("h:mm")
    .aFmts(22)=("h:mm:ss")
    .aFmts(23)=("m/d/yy h:mm")
    .aFmts(24)=("(#,##0_);(#,##0)")
    .aFmts(25)=("(#,##0_);[Red](#,##0)")
    .aFmts(26)=("(#,##0.00_);(#,##0.00)")
    .aFmts(27)=("(#,##0.00_);[Red](#,##0.00)")
    .aFmts(28)=('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)')
    .aFmts(29)=('_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)')
    .aFmts(30)=('_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)')
    .aFmts(31)=('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')
    .aFmts(32)=("mm:ss")
    .aFmts(33)=("[h]:mm:ss")
    .aFmts(34)=("mm:ss.0")
    .aFmts(35)=("##0.0E+0")
    .aFmts(36)=("@")
    .aFmts(37)=("dd/mm/yyyy")
    .aFmts(38)=("$#,##0.00")
    
	ENDWITH 
ENDFUNC 


FUNCTION WriteFormatTable
	* Creating the table array
	LOCAL lcBiffFormatRecord
	lcBiffFormatRecord = ""

	FOR n = 1 TO ALEN(This.aFmts,1)
		lcBiffFormatRecord = lcBiffFormatRecord + This.WriteFormatRecord(This.aFmts(n))
	ENDFOR
	RETURN lcBiffFormatRecord
ENDFUNC 


FUNCTION WriteFormatRecord(tcFormat)
	* #DEFINE FormatRecord         0x001E
	LOCAL lcReturn
	lcReturn = BINTOC(FormatRecord, "2RS") + ;
			BINTOC(LEN(tcFormat) + 1,"2RS") + ;
			CHR(LEN(tcFormat)) + ;
			tcFormat
	RETURN lcReturn 
ENDFUNC 


FUNCTION WriteWindowProtectRecord
	* #DEFINE WindowProtectRecord  0x0019
	LOCAL lcBiffWindowProtectRecord
	lcBiffWindowProtectRecord = BINTOC(WindowProtectRecord, "2RS")
	RETURN lcBiffWindowProtectRecord
ENDFUNC 


FUNCTION WriteXFTable
	* #DEFINE XFRecord             0x0243
	LOCAL lcBiffXFTable, lcAuxTable, lcArrayTable, lnItem, n

	lcAuxTable = ;
	     "0x0243,0x00C,0x0000,0x03F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0001,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0001,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0002,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0002,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0000,0x0001,0x0000,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x2101,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x1F01,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x2001,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
    	 "0x0243,0x00C,0x1E01,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
	     "0x0243,0x00C,0x0901,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000"

	lcArrayTable = 0h
	FOR n = 1 TO (8*21)
		lnItem = EVALUATE(GETWORDNUM(lcAuxTable, n, ","))
		lcArrayTable = 	lcArrayTable + LEFT(BINTOC(lnItem, "4RS"),2)
	ENDFOR

	lcBiffXFTable = BINTOC(2, "4RS") + ;
					lcArrayTable				

	lcBiffFX = This.CreateFxTable()
	RETURN lcBiffXFTable + lcBiffFx
ENDFUNC 


FUNCTION CreateFxTable
	LOCAL n, lnCells, lcTable
	lcTable = ""
	lnCells = ALEN(This.aCells, 1)
	FOR n = 1 TO lnCells
		lcTable = lcTable + This.WriteFXRecord(n)
	ENDFOR 
	RETURN lcTable
ENDFUNC 



FUNCTION WriteFXRecord(tnCell)
* #DEFINE ExtendedRecord       0x0243

LOCAL lcRecord, lnFontIndex, lnFormatIndex, lnHorizontalAlignment, lnForeColor, lnBackColor
lnFontIndex           = EVL(This.aCells(tnCell, 4), 0)
lnFormatIndex         = EVL(This.aCells(tnCell, 5), 0)
lnHorizontalAlignment = EVL(This.aCells(tnCell, 6), 0) && 2 = Center
lnBackColor           = EVL(This.aCells(tnCell, 8), 0)

IF lnFontIndex = 0 AND lnFormatIndex = 0
	This.aCells(tnCell,10) = 0
	RETURN ""
ENDIF 

This._nCurrFXRecord = This._nCurrFXRecord + 1
This.aCells(tnCell,10) = This._nCurrFXRecord

LOCAL lnAttr
lnAttr = 0

IF lnFontIndex > 0
	lnAttr = BITOR(lnAttr, 2)
ENDIF 
IF lnHorizontalAlignment > 0 && Alignment.General
	lnAttr = BITOR(lnAttr, 4)
ENDIF 
IF lnBackColor > 0 && Alignment.General
	lnAttr = BITOR(lnAttr, 0x10)
ENDIF 
lnAttr = BITLSHIFT(lnAttr, 2)

lcRecord = 0h + BINTOC(ExtendedRecord, "2RS") + ;
			BINTOC(0x00C, "2RS") + ;
			CHR(lnFontIndex) + ; && FontIndex
			CHR(lnFormatIndex) + ; && FormatIndex
			CHR(1) + ;
			CHR(lnAttr)
LOCAL lqBackColor, lnXlsColor
IF (EMPTY(lnBackColor)) OR (VARTYPE(lnBackColor) <> "N")
	lqBackColor = 0h00CE && 0xCE
ELSE 
	lnXlsColor = This.aColors(lnBackColor,3)
	lqBackColor = 0h + LEFT(BINTOC(lnXLSColor, "4RS"),2)
ENDIF 

lcRecord = lcRecord + BINTOC(lnHorizontalAlignment, "2RS") + ;
				lqBackColor + ;
				BINTOC(0x0000, "2RS") + ;
				BINTOC(0x0000, "2RS")

RETURN lcRecord
ENDFUNC 



FUNCTION WriteStyleTable
	* #DEFINE StyleRecord          0x0293   Linha 4D0
	LOCAL lcBiffStyle

	lcBiffStyle = ;
		0h + BINTOC(StyleRecord, "2RS") + 0h0400 + 0h108003FF + ;
		0h + BINTOC(StyleRecord, "2RS") + 0h0C00 + 0h110009436F6D6D61205B305D + ;
		0h + BINTOC(StyleRecord, "2RS") + 0h0400 + 0h128004FF + ;	
		0h + BINTOC(StyleRecord, "2RS") + 0h0F00 + 0h13000C43757272656E6379205B305D + ;
		0h + BINTOC(StyleRecord, "2RS") + 0h0400 + 0h008000FF + ;
		0h + BINTOC(StyleRecord, "2RS") + 0h0400 + 0h148005FF

	RETURN lcBiffStyle
ENDFUNC 


FUNCTION WriteDefaultRowHeight
	* #DEFINE DefRowHeight          0x0225
	LOCAL lqReturn
	lqReturn = 0h
	IF NOT EMPTY(This.nDefaultRowHeight)
	    lqReturn = 0h + ;
	    	BINTOC(DefRowHeight, "2RS") + ;
    		BINTOC(4, "2RS") + ;
    		BINTOC(1, "2RS") + ;
    		BINTOC(This.nDefaultRowHeight * 20,"2RS")
	ENDIF 
	RETURN lqReturn

*!*	5.31 DEFAULTROWHEIGHT
*!*	BIFF2 BIFF3 BIFF4 BIFF5 BIFF8
*!*	0025H 0225H 0225H 0225H 0225H
*!*	This record specifies the default height and default flags for rows that do not have a corresponding ROW record
*!*	(?5.88).
*!*	Record DEFAULTROWHEIGHT, BIFF2:
*!*	Offset Size Contents
*!*	0 2 Default height for unused rows, in twips = 1/20 of a point
*!*	Bit Mask Contents
*!*	0-14 7FFFH Default height for unused rows, in twips = 1/20 of a point
*!*	15 8000H 1 = Row height not changed manually
*!*	Record DEFAULTROWHEIGHT, BIFF3-BIFF8:
*!*	Offset Size Contents
*!*	0 2 Option flags:
*!*	          Bit Mask Contents
*!*	          0 0001H 1 = Row height and default font height do not match
*!*	          1 0002H 1 = Row is hidden
*!*	          2 0004H 1 = Additional space above the row
*!*	          3 0008H 1 = Additional space below the row
*!*	2 2 Default height for unused rows, in twips = 1 /20 of a point
ENDFUNC 



FUNCTION WriteDefaultColumnWidth
	* #DEFINE DefColWidth          0x0055
	LOCAL lqReturn
	lqReturn = 0h
	IF NOT EMPTY(This.nDefaultColumnWidth)
	    lqReturn = 0h + ;
	    	BINTOC(DefColWidth, "2RS") + ;
    		BINTOC(2, "2RS") + ;
    		BINTOC(This.nDefaultColumnWidth,"2RS")
	ENDIF 
	RETURN lqReturn

*!*	5.32 DEFCOLWIDTH
*!*	BIFF2 BIFF3 BIFF4 BIFF5 BIFF8
*!*	0055H 0055H 0055H 0055H 0055H
*!*	This record specifies the default column width for columns that do not have a specific width set using the records
*!*	COLWIDTH (BIFF2, ?5.20), COLINFO (BIFF3-BIFF8, ?5.18), or STANDARDWIDTH (?5.101).
*!*	Record DEFCOLWIDTH, BIFF2-BIFF8:
*!*	Offset Size Contents
*!*	0 2 Column width in characters, using the width of the zero character from default font (first
*!*	FONT record in the file). Excel adds some extra space to the default width, depending on
*!*	the default font and default font size. The algorithm how to exactly calculate the resulting
*!*	column width is not known.
*!*	Example: The default width of 8 set in this record results in a column width of
*!*	8.43 using Arial font with a size of 10 points.

ENDFUNC 




FUNCTION SetColumnWidth
	LPARAMETERS tnColumn, tnWidth
	LOCAL lnCols, lnNextCol
	lnCols = ALEN(This.aColWidths, 1)
	IF VARTYPE(This.aColWidths(1)) <> "N"
		lnNextCol = 1
	ELSE 
		lnNextCol = lnCols + 1
	ENDIF
	DIMENSION This.aColWidths(lnNextCol,2)
	This.aColWidths(lnNextCol, 1) = tnColumn
	This.aColWidths(lnNextCol, 2) = tnWidth
ENDFUNC 


FUNCTION WriteColumnInfoRecord
	LPARAMETERS tnCol, tnWidth
	* #DEFINE ColumnInfoRecord     0x007D

	tnCol = MAX(0, tnCol - 1)
    
    LOCAL lcReturn
    lcReturn = 0h + ;
    	BINTOC(ColumnInfoRecord, "2RS") + ;
    		BINTOC(12, "2RS") + ;
    		BINTOC(tnCol, "2RS") + ;
    		BINTOC(tnCol, "2RS") + ;
    		BINTOC((tnWidth * 256 / 7),"2RS") + ;
    		BINTOC(15, "2RS") + ;
    		BINTOC(0, "2RS") + ;
    		BINTOC(0, "2RS")
	RETURN lcReturn
ENDFUNC


FUNCTION WriteCell(tnIndex, tnRow, tnCol, tuValue)
	LOCAL lcType
	lcType = VARTYPE(tuValue)

	DO CASE
	CASE lcType = "C"
		This.WriteString(tnIndex, tnRow, tnCol, tuValue)

	CASE lcType = "N"
		This.WriteDoble(tnIndex, tnRow, tnCol, tuValue)

	CASE lcType = "D"
		This.WriteDate(tnIndex, tnRow, tnCol, tuValue)

	OTHERWISE
		This.WriteEmpty(tnIndex, tnRow, tnCol)

	ENDCASE
RETURN 



FUNCTION WriteString(tnIndex, tnRow, tnCol, tcString)
	IF LEN(tcString) > 255
		tcString = LEFT(tcString, 255)
	ENDIF 

	LOCAL lnFXIndex
	lnFXIndex = EVL(This.aCells(tnIndex, 10), 0)

	* #DEFINE LabelRecord          0x0204
	This.cXLS = This.cXLS + ;
		BINTOC(LabelRecord,"2RS") + BINTOC(8+LEN(tcString),"2RS") + ;
		BINTOC(tnRow - 1,"2RS") + BINTOC(tnCol - 1,"2RS") + BINTOC(lnFXIndex,"2RS") + BINTOC(LEN(tcString),"2RS") + tcString
	RETURN 
RETURN 
	

FUNCTION WriteInt(tnIndex, tnRow, tnCol, tnValue)	
	LOCAL lnValue
	lnValue = BITLSHIFT(tnValue,2) + 2
	LOCAL lnFXIndex
	lnFXIndex = This.aCells(tnIndex, 10)
	This.cXLS = This.cXLS + ;
		BINTOC(0x27e,"2RS") + BINTOC(10,"2RS") + ;
		BINTOC(tnRow - 1,"2RS") + BINTOC(tnCol - 1,"2RS") + BINTOC(lnFXIndex,"2RS") + BINTOC(lnValue,"4RS")
	RETURN 


FUNCTION WriteDoble(tnIndex, tnRow, tnCol, tnValue)	
	* #DEFINE NumberRecord         0x0203
	LOCAL lnFXIndex
	lnFXIndex = This.aCells(tnIndex, 10)
	This.cXLS = This.cXLS + ;
		BINTOC(NumberRecord,"2RS") + BINTOC(14,"2RS")     + ; && Headers
		BINTOC(tnRow - 1,"2RS") + BINTOC(tnCol - 1,"2RS") + ; && Row, Col
		BINTOC(lnFXIndex, "2RS")                           + ; && (ushort)cell.FXIndex
		BINTOC(tnValue,"8SB")
RETURN


FUNCTION WriteDate(tnIndex, tnRow, tnCol, tdDate as Date)
	* #DEFINE NumberRecord         0x0203
	LOCAL lnDateValue, lcDateFormat, lnFXIndex, lcReturn
	lnDateValue = tdDate - {^1900-01-01} + 2
	lnFXIndex = This.aCells(tnIndex, 10)
	lcReturn = 0h + ;
		BINTOC(NumberRecord,"2RS") + BINTOC(14,"2RS") + BINTOC(tnRow - 1,"2RS") + BINTOC(tnCol - 1,"2RS") + ;
			BINTOC(lnFXIndex, "2RS") + ; &&(ushort)cell.FXIndex }
			BINTOC(lnDateValue,"8SB")
	This.cXLS = This.cXLS + ;
		lcReturn
	RETURN



FUNCTION WriteEmpty(tnIndex, tnRow, tnCol)
	LOCAL lnFXIndex
	lnFXIndex = This.aCells(tnIndex, 10)
	This.cXLS = This.cXLS + ;
		BINTOC(0x0201,"2RS") + BINTOC(6,"2RS")            + ;  && Header
		BINTOC(tnRow - 1,"2RS") + BINTOC(tnCol - 1,"2RS") + ; && Row, Col
		BINTOC(lnFXIndex,"2RS")
	RETURN 

ENDDEFINE