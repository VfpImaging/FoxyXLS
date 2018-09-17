* FOXYXLS Class
* Creates pure XLS files without the need of having MS-OFFICE installed
* Author: VFPIMAGING
* Based on Serhiy Perevoznyk's work published in his blog
* http://delphi32.blogspot.com.br/2011/06/generate-excel-files-without-using.html
*
* May 2014 - Victor Espina
* Changes to make it VFP6 compatible:
* 1) Change ASCAN() function for a custom implementation ASCANX()
* 2) Custom implementation of EVL() and GETWORDNUM() functions
* 3) Change of inline binary values (0h00) for a more generic function (S2B("0h00"))
* 4) Use of a TRY-CATCH substitute (VFP6 only)
* 5) Change BINTOC() function for a custom implementation BINTOCX()
*

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

DEFINE CLASS FoxyXLS AS CUSTOM
	m.cXLS				  = ""
	m.cFile				  = ""
	m.cDateType			  = "B"
	m.cHeader			  = ""
	m.cFooter			  = ""
	m.cAuthor			  = ""
	m.nCodePage			  = 1252
	m._nCurrFXRecord	  = 20
	m.cDefaultFontName	  = "Arial"
	m.nDefaultFontSize	  = 10
	m.nDefaultColumnWidth = 0
	m.nDefaultRowHeight	  = 0

	FUNCTION INIT

		THIS.ADDPROPERTY("aFonts[1,1]", "")
		THIS.ADDPROPERTY("aCells[1,1]", "")
		THIS.ADDPROPERTY("aFmts[1,1]", "")
		THIS.ADDPROPERTY("aColWidths[1,1]", "")
		THIS.ADDPROPERTY("aColors[1,1]", "")

		THIS.PrepareFormatArray()
		THIS.PrepareColorsArray()

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
		THIS.cDateType = lcDateType

		* Prepare the fonts Array
		DIMENSION THIS.aFonts(4, 2)
		THIS.aFonts(1, 1) = "Arial,10,"
		THIS.aFonts(2, 1) = "Arial,10,B"
		THIS.aFonts(3, 1) = "Arial,10,I"
		THIS.aFonts(4, 1) = "Arial,10,BI"

		THIS.aFonts(1, 2) = 0
		THIS.aFonts(2, 2) = 0
		THIS.aFonts(3, 2) = 0
		THIS.aFonts(4, 2) = 0
	ENDFUNC






	FUNCTION WriteFile(tcFile)
		IF EMPTY(tcFile)
			tcFile = THIS.cFile
		ENDIF
		THIS.CreateHeader()

		LOCAL N, lcBiffCol
		lcBiffCol = ""


		* Setting the default row heights
		lcBiffCol = lcBiffCol + THIS.WriteDefaultRowHeight()


		* Setting the default row heights
		*	lcBiffCol = lcBiffCol + ;
		0h0802 + ;  && Index
		0h1000 + ;  && Size
		0h0000 + ;  && Row number
		0h0000 + ;  && Index to column of the first cell which is described by a cell record
		0h0400 + ;  && Index to column of the last cell which is described by a cell record, increased by 1
		BINTOCX(40 * 20, "2RS") + ;  && Height of the row, in twips = 1/20 of a point
		0h0000 + ;  && Not used
		0h1022 + ;  && a relative offset to calculate stream position of the first cell record for this row (?4.7.2)
		0h0100f000  && Option flags and default row formatting

		* Setting the default column widths
		lcBiffCol = lcBiffCol + THIS.WriteDefaultColumnWidth()


		* Setting the column forced widths
		IF VARTYPE(THIS.aColWidths(1, 1)) = "N"
			FOR N = 1 TO ALEN(THIS.aColWidths, 1)
				lcBiffCol = lcBiffCol + THIS.WriteColumnInfoRecord(THIS.aColWidths(N, 1), THIS.aColWidths(N, 2))
			ENDFOR
		ENDIF

		* Writing the cells
		FOR N = 1 TO ALEN(THIS.aCells, 1)
			THIS.WriteCell(N, THIS.aCells(N, 1), THIS.aCells(N, 2), THIS.aCells(N, 3))
		ENDFOR

		THIS.cXLS = THIS.cHeader + lcBiffCol + THIS.cXLS + THIS.cFooter
		= STRTOFILE(THIS.cXLS, tcFile)
	ENDFUNC



	PROCEDURE GetFontFullName
		LPARAMETERS tnIndex, tcFontFull, tnForeColor, tcColorName, tnXLSColor

		* Check if only the forecolor was passed
		LOCAL lnIndex
		*:Global aColors[1]
		IF EMPTY(tcFontFull) AND tnForeColor > 0
			#IF VERSION(5) > 600
				TRY
				#ELSE
					TRY()
					#ENDIF
					tcColorName	= THIS.aColors(tnForeColor, 1)
					tnXLSColor	= THIS.aColors(tnForeColor, 2)
					lcFontFull	= ",,," + tcColorName
					#IF VERSION(5) > 600
					CATCH
					ENDTRY
				#ELSE
				CATCH()
				ENDTRY()
			#ENDIF
			RETURN lcFontFull
		ENDIF


		* Working with the supposition NOT EMPTY(tcFontFull)

		* Getting the correct Forecolor value
		* And preparing the internal tables to work with it
		LOCAL lcFontFull, lcColorName, lnColor
		LOCAL lcFont, lnSize, lcStyle, lcColor
		lcFont  = ALLTRIM(GETWORDNUM(tcFontFull, 1, ","))
		IF EMPTY(lcFont)
			lcFont = THIS.cDefaultFontName
		ENDIF
		lnSize  = INT(VAL(ALLTRIM(GETWORDNUM(tcFontFull, 2, ","))))
		IF EMPTY(lnSize)
			lnSize = THIS.nDefaultFontSize
		ENDIF
		lcStyle	= UPPER(ALLTRIM(GETWORDNUM(tcFontFull, 3, ",")))
		lcColor	= UPPER(ALLTRIM(GETWORDNUM(tcFontFull, 4, ",")))

		tcFontFull = lcFont + "," + TRANSFORM(lnSize) + "," + lcStyle + ","

		DO CASE
			CASE EMPTY(lcColor) AND tnForeColor = 0 && No color passed
				tcColorName	= ""
				tnXLSColor	= 0
				lcFontFull	= tcFontFull

			CASE NOT EMPTY(lcColor)
				#IF VERSION(5) > 600
					TRY
					#ELSE
						TRY()
						#ENDIF
						LOCAL ARRAY aColors[1]
						ACOPY(THIS.aColors, aColors)
						lnIndex = CEILING(ASCANX(@aColors, lcColor, 1, 0, 1, 4 + 2 + 1) / 3)
						IF lnIndex = 0 && Not found, the ignore the color
							tcColorName	= ""
							tnXLSColor	= 0
							lcFontFull	= tcFontFull
						ELSE
							&& Found the color
							tcColorName	= lcColor
							tnXLSColor	= THIS.aColors(lnIndex, 2)
							lcFontFull	= tcFontFull + lcColor
						ENDIF
						#IF VERSION(5) > 600
						CATCH
							tcColorName	= ""
							tnXLSColor	= 0
							lcFontFull	= tcFontFull
						ENDTRY
					#ELSE
						IF CATCH()
							tcColorName	= ""
							tnXLSColor	= 0
							lcFontFull	= tcFontFull
						ENDIF
					ENDTRY()
				#ENDIF



			OTHERWISE
				tcColorName	= ""
				tnXLSColor	= 0
				lcFontFull	= tcFontFull
		ENDCASE

		RETURN lcFontFull
	ENDPROC



	PROCEDURE AddCell(tnRow, tnCol, tuValue, tcFontFull, tcFormat, tnAlign, tnForeColor, tnBackColor)
		LOCAL lnFontIndex, lnFontTotal
		tcFontFull	= EVL(tcFontFull, "")
		tnForeColor	= EVL(tnForeColor, 0)
		* Prepare the Cells array, having the next record initialized to receive the information
		LOCAL lnTotCells, i
		lnTotCells = ALEN(THIS.aCells, 1)
		IF lnTotCells = 1 AND VARTYPE(THIS.aCells(1)) <> "N"
			i = 1
		ELSE
			i = lnTotCells + 1
		ENDIF
		DIMENSION THIS.aCells(i, 10)

		LOCAL lcFontFull, lnXLSColor, lcColorName
		lcFontFull	= ""
		lnXLSColor	= 0
		lcColorName	= ""
		IF EMPTY(tcFontFull) AND EMPTY(tnForeColor)
		ELSE
			lcFontFull = THIS.GetFontFullName(i, tcFontFull, tnForeColor, @lcColorName, @lnXLSColor)
		ENDIF

		IF EMPTY(lcFontFull)
			lnFontIndex = 0
		ELSE
			LOCAL ARRAY aFonts[1]
			ACOPY(THIS.aFonts, aFonts)
			lnFontIndex = CEILING(ASCANX(@aFonts, lcFontFull, 1, 0, 1, 4 + 2 + 1) / 2)
			IF lnFontIndex = 0
				lnFontTotal	= ALEN(THIS.aFonts, 1)
				lnFontIndex	= lnFontTotal + 1
				DIMENSION THIS.aFonts(lnFontIndex, 2)
				THIS.aFonts(lnFontIndex, 1)	= lcFontFull
				THIS.aFonts(lnFontIndex, 2)	= lnXLSColor
			ENDIF
		ENDIF

		LOCAL lcValueType
		lcValueType = VARTYPE(m.tuValue)
		IF lcValueType = "D" AND EMPTY(tcFormat)
			DO CASE
				CASE THIS.cDateType = "A"
					tcFormat = "m/d/yy"
				CASE THIS.cDateType = "B"
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
				LOCAL ARRAY aFmts[1]
				ACOPY(THIS.aFmts, aFmts)
				lnFormatIndex = ASCANX(@aFmts, ALLTRIM(luFmtIndex), 1, 0, 1, 4 + 2 + 1) && All records, column 1, Case insensitive
				lnFormatIndex = MAX(lnFormatIndex - 1, 0)
			OTHERWISE
				lnFormatIndex = 0
		ENDCASE

		THIS.aCells(i, 1) = tnRow
		THIS.aCells(i, 2) = tnCol
		THIS.aCells(i, 3) = tuValue
		THIS.aCells(i, 4) = lnFontIndex
		THIS.aCells(i, 5) = lnFormatIndex
		THIS.aCells(i, 6) = tnAlign
		THIS.aCells(i, 7) = lnXLSColor
		THIS.aCells(i, 8) = tnBackColor
	ENDPROC


	FUNCTION CreateHeader
		LOCAL lcBiffFooter, lcBiffHeader
		THIS.cFooter = BINTOCX(EOFRecord, "4RS")

		LOCAL lcBiffStart, lcBiffAuthor, lcBiffCodePage, lcBiffFontTable, lcBiffHeaderRecord, lcBiffFooterRecord
		LOCAL lcBiffFormatTable, lcBiffWindowProtect, lcBiffXFTable, lcBiffStyle
		lcBiffStart			= THIS.GetMainHeader()
		lcBiffAuthor		= THIS.WriteAuthorRecord()
		lcBiffCodePage		= THIS.WriteCodepageRecord()
		lcBiffFontTable		= THIS.WriteFontTable()
		lcBiffHeader		= THIS.WriteHeaderRecord()
		lcBiffFooter		= THIS.WriteFooterRecord()
		lcBiffFormatTable	= THIS.WriteFormatTable()
		lcBiffWindowProtect	= THIS.WriteWindowProtectRecord()
		lcBiffXFTable		= THIS.WriteXFTable()
		lcBiffStyle			= THIS.WriteStyleTable()

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

		THIS.cHeader = lcFullHeader
	ENDFUNC


	FUNCTION GetMainHeader
		LOCAL lcBiffHeader
		lcBiffHeader = S2B("0h090208000000100000000000")
		RETURN lcBiffHeader
	ENDFUNC

	FUNCTION WriteAuthorRecord
		LOCAL lcBiffAuthor
		lcBiffAuthor = S2B("0h5C00200006") + PADR(THIS.cAuthor, 31, " ")  && 31 positions
		RETURN lcBiffAuthor
	ENDFUNC

	FUNCTION WriteCodepageRecord
		LOCAL lcBiffCodePage
		IF VARTYPE(THIS.nCodePage) = "C"
			THIS.nCodePage = VAL(THIS.nCodePage)
		ENDIF

		lcBiffCodePage = S2B("0h42000200") + S2B("0h") + BINTOCX(THIS.nCodePage, "2RS")
		RETURN lcBiffCodePage
	ENDFUNC


	FUNCTION WriteFontTable
		* Write font table
		LOCAL lcBiffFontRecord, N, lcFont
		lcBiffFontRecord = ""
		FOR N = 1 TO ALEN(THIS.aFonts, 1)
			lcFont = THIS.aFonts(N, 1)
			IF NOT EMPTY(lcFont)
				lcBiffFontRecord = lcBiffFontRecord + THIS.WriteFontRecord(N, lcFont)
			ENDIF
		ENDFOR
		RETURN lcBiffFontRecord
	ENDFUNC


	FUNCTION WriteFontRecord(tnIndex, tcFontFull)

		* Getting the correct Forecolor value
		LOCAL lcFont, lnSize, lcStyle, lcArray, lnIntStyle, lcColor, lnColor, lcColorName, lnClrIndex
		lcFont	= ALLTRIM(GETWORDNUM(tcFontFull, 1, ","))
		lnSize	= VAL(ALLTRIM(GETWORDNUM(tcFontFull, 2, ",")))
		lcStyle	= UPPER(ALLTRIM(GETWORDNUM(tcFontFull, 3, ",")))
		lnColor	= THIS.aFonts(tnIndex, 2)

		lnIntStyle = IIF("B" $ lcStyle, 1, 0) + ;
			IIF("I" $ lcStyle, 2, 0) + ;
			IIF("U" $ lcStyle, 4, 0) + ;
			IIF("S" $ lcStyle, 8, 0)

		lcArray = BINTOCX(FontRecord, "2RS")   + ; && 0h3102 = FontRecord
			BINTOCX(LEN(lcFont) + 7, "2RS") + ; && Fontname length + 7
			BINTOCX(lnSize * 20, "2RS")    + ; && FontSize
			BINTOCX(lnIntStyle, "2RS")     + ; && FontStyle
			BINTOCX(lnColor, "2RS")   + ; && Color
			CHR(LEN(lcFont))              + ;
			lcFont

		RETURN lcArray
	ENDFUNC


	FUNCTION PrepareColorsArray
		DIMENSION THIS.aColors(17, 3)
		WITH THIS
			.aColors(1, 1) = "WHITE"
			.aColors(1, 2) = 0x01    && ForeColor
			.aColors(1, 3) = 0xC041  && BackColor

			.aColors(2, 1) = "RED"
			.aColors(2, 2) = 0x02    && ForeColor
			.aColors(2, 3) = 0xC081  && BackColor

			.aColors(3, 1) = "GREEN"
			.aColors(3, 2) = 0x03    && ForeColor
			.aColors(3, 3) = 0xC0C1  && BackColor

			.aColors(4, 1) = "BLUE"
			.aColors(4, 2) = 0x04    && ForeColor
			.aColors(4, 3) = 0xC101  && BackColor

			.aColors(5, 1) = "YELLOW"
			.aColors(5, 2) = 0x05    && ForeColor
			.aColors(5, 3) = 0xC141  && BackColor

			.aColors(6, 1) = "MAGENTA"
			.aColors(6, 2) = 0x06    && ForeColor
			.aColors(6, 3) = 0xC181  && BackColor

			.aColors(7, 1) = "CYAN"
			.aColors(7, 2) = 0x07    && ForeColor
			.aColors(7, 3) = 0xC1C1  && BackColor

			.aColors(8, 1) = "DARKRED"
			.aColors(8, 2) = 0x10    && ForeColor
			.aColors(8, 3) = 0xC401  && BackColor

			.aColors(9, 1) = "DARKGREEN"
			.aColors(9, 2) = 0x11    && ForeColor
			.aColors(9, 3) = 0xC441  && BackColor

			.aColors(10, 1)	= "DARKBLUE"
			.aColors(10, 2)	= 0x12    && ForeColor
			.aColors(10, 3)	= 0xC481  && BackColor

			.aColors(11, 1)	= "OLIVE"
			.aColors(11, 2)	= 0x13    && ForeColor
			.aColors(11, 3)	= 0xC4C1  && BackColor

			.aColors(12, 1)	= "PURPLE"
			.aColors(12, 2)	= 0x14    && ForeColor
			.aColors(12, 3)	= 0xC501  && BackColor

			.aColors(13, 1)	= "TEAL"
			.aColors(13, 2)	= 0x15    && ForeColor
			.aColors(13, 3)	= 0xC541  && BackColor

			.aColors(14, 1)	= "SILVER"
			.aColors(14, 2)	= 0x16    && ForeColor
			.aColors(14, 3)	= 0xC581  && BackColor

			.aColors(15, 1)	= "GRAY"
			.aColors(15, 2)	= 0x17    && ForeColor
			.aColors(15, 3)	= 0xC5C1  && BackColor

			.aColors(16, 1)	= "BLACK"
			.aColors(16, 2)	= 0x00    && ForeColor
			.aColors(16, 3)	= 0xC001  && BackColor

			.aColors(17, 1)	= "AUTOMATIC"
			.aColors(17, 2)	= 0x7fff  && ForeColor
			.aColors(17, 3)	= 0xCE00  && BackColor
		ENDWITH
	ENDFUNC



	FUNCTION WriteHeaderRecord
		LOCAL lcBiffHeaderRecord
		lcBiffHeaderRecord = BINTOCX(HeaderRecord, "4RS")
		RETURN lcBiffHeaderRecord
	ENDFUNC


	FUNCTION WriteFooterRecord
		LOCAL lcBiffFooterRecord
		lcBiffFooterRecord = BINTOCX(FooterRecord, "4RS")
		RETURN lcBiffFooterRecord
	ENDFUNC


	FUNCTION PrepareFormatArray
		* Prepare the Formatting record
		DIMENSION THIS.aFmts(38)
		WITH THIS
			.aFmts(1)  = ("General")
			.aFmts(2)  = ("0")
			.aFmts(3)  = ("0.00")
			.aFmts(4)  = ("#,##0")
			.aFmts(5)  = ("#,##0.00")
			.aFmts(6)  = ("($#,##0_);($#,##0)")
			.aFmts(7)  = ("($#,##0_);[Red]($#,##0)")
			.aFmts(8)  = ("($#,##0.00_);($#,##0.00)")
			.aFmts(9)  = ("($#,##0.00_);[Red]($#,##0.00)")
			.aFmts(10) = ("0%")
			.aFmts(11) = ("0.00%")
			.aFmts(12) = ("0.00E+00")
			.aFmts(13) = ("# ?/?")
			.aFmts(14) = ("# ??/??")
			.aFmts(15) = ("m/d/yy")
			.aFmts(16) = ("d-mmm-yy")
			.aFmts(17) = ("d-mmm")
			.aFmts(18) = ("mmm-yy")
			.aFmts(19) = ("h:mm AM/PM")
			.aFmts(20) = ("h:mm:ss AM/PM")
			.aFmts(21) = ("h:mm")
			.aFmts(22) = ("h:mm:ss")
			.aFmts(23) = ("m/d/yy h:mm")
			.aFmts(24) = ("(#,##0_);(#,##0)")
			.aFmts(25) = ("(#,##0_);[Red](#,##0)")
			.aFmts(26) = ("(#,##0.00_);(#,##0.00)")
			.aFmts(27) = ("(#,##0.00_);[Red](#,##0.00)")
			.aFmts(28) = ('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)')
			.aFmts(29) = ('_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)')
			.aFmts(30) = ('_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)')
			.aFmts(31) = ('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')
			.aFmts(32) = ("mm:ss")
			.aFmts(33) = ("[h]:mm:ss")
			.aFmts(34) = ("mm:ss.0")
			.aFmts(35) = ("##0.0E+0")
			.aFmts(36) = ("@")
			.aFmts(37) = ("dd/mm/yyyy")
			.aFmts(38) = ("$#,##0.00")

		ENDWITH
	ENDFUNC


	FUNCTION WriteFormatTable
		* Creating the table array
		LOCAL lcBiffFormatRecord
		*:Global n
		lcBiffFormatRecord = ""

		FOR N = 1 TO ALEN(THIS.aFmts, 1)
			lcBiffFormatRecord = lcBiffFormatRecord + THIS.WriteFormatRecord(THIS.aFmts(N))
		ENDFOR
		RETURN lcBiffFormatRecord
	ENDFUNC


	FUNCTION WriteFormatRecord(tcFormat)
		* #DEFINE FormatRecord         0x001E
		LOCAL lcReturn
		lcReturn = BINTOCX(FormatRecord, "2RS") + ;
			BINTOCX(LEN(tcFormat) + 1, "2RS") + ;
			CHR(LEN(tcFormat)) + ;
			tcFormat
		RETURN lcReturn
	ENDFUNC


	FUNCTION WriteWindowProtectRecord
		* #DEFINE WindowProtectRecord  0x0019
		LOCAL lcBiffWindowProtectRecord
		lcBiffWindowProtectRecord = BINTOCX(WindowProtectRecord, "2RS")
		RETURN lcBiffWindowProtectRecord
	ENDFUNC


	FUNCTION WriteXFTable
		* #DEFINE XFRecord             0x0243
		LOCAL lcBiffXFTable, lcAuxTable, lcArrayTable, lnItem, N
		LOCAL lcBiffFX

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

		lcArrayTable = S2B("0h")
		FOR N = 1 TO (8 * 21)
			lnItem		 = EVALUATE(GETWORDNUM(lcAuxTable, N, ","))
			lcArrayTable = 	lcArrayTable + LEFT(BINTOCX(lnItem, "4RS"), 2)
		ENDFOR

		lcBiffXFTable = BINTOCX(2, "4RS") + ;
			lcArrayTable

		lcBiffFX = THIS.CreateFxTable()
		RETURN lcBiffXFTable + lcBiffFX
	ENDFUNC


	FUNCTION CreateFxTable
		LOCAL N, lnCells, lcTable
		lcTable	= ""
		lnCells	= ALEN(THIS.aCells, 1)
		FOR N = 1 TO lnCells
			lcTable = lcTable + THIS.WriteFXRecord(N)
		ENDFOR
		RETURN lcTable
	ENDFUNC



	FUNCTION WriteFXRecord(tnCell)
		* #DEFINE ExtendedRecord       0x0243

		LOCAL lcRecord, lnFontIndex, lnFormatIndex, lnHorizontalAlignment, lnForeColor, lnBackColor
		lnFontIndex			  = EVL(THIS.aCells(tnCell, 4), 0)
		lnFormatIndex		  = EVL(THIS.aCells(tnCell, 5), 0)
		lnHorizontalAlignment = EVL(THIS.aCells(tnCell, 6), 0) && 2 = Center
		lnBackColor			  = EVL(THIS.aCells(tnCell, 8), 0)

		IF lnFontIndex = 0 AND lnFormatIndex = 0
			THIS.aCells(tnCell, 10) = 0
			RETURN ""
		ENDIF

		THIS._nCurrFXRecord		= THIS._nCurrFXRecord + 1
		THIS.aCells(tnCell, 10)	= THIS._nCurrFXRecord

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

		lcRecord = S2B("0h") + BINTOCX(ExtendedRecord, "2RS") + ;
			BINTOCX(0x00C, "2RS") + ;
			CHR(lnFontIndex) + ; && FontIndex
			CHR(lnFormatIndex) + ; && FormatIndex
			CHR(1) + ;
			CHR(lnAttr)
		LOCAL lqBackColor, lnXLSColor
		IF (EMPTY(lnBackColor)) OR (VARTYPE(lnBackColor) <> "N")
			lqBackColor = S2B("0h00CE") && 0xCE
		ELSE
			lnXLSColor	= THIS.aColors(lnBackColor, 3)
			lqBackColor	= S2B("0h") + LEFT(BINTOCX(lnXLSColor, "4RS"), 2)
		ENDIF

		lcRecord = lcRecord + BINTOCX(lnHorizontalAlignment, "2RS") + ;
			lqBackColor + ;
			BINTOCX(0x0000, "2RS") + ;
			BINTOCX(0x0000, "2RS")

		RETURN lcRecord
	ENDFUNC



	FUNCTION WriteStyleTable
		* #DEFINE StyleRecord          0x0293   Linha 4D0
		LOCAL lcBiffStyle

		lcBiffStyle = ;
			S2B("0h") + BINTOCX(StyleRecord, "2RS") + S2B("0h0400") + S2B("0h108003FF") + ;
			S2B("0h") + BINTOCX(StyleRecord, "2RS") + S2B("0h0C00") + S2B("0h110009436F6D6D61205B305D") + ;
			S2B("0h") + BINTOCX(StyleRecord, "2RS") + S2B("0h0400") + S2B("0h128004FF") + ;
			S2B("0h") + BINTOCX(StyleRecord, "2RS") + S2B("0h0F00") + S2B("0h13000C43757272656E6379205B305D") + ;
			S2B("0h") + BINTOCX(StyleRecord, "2RS") + S2B("0h0400") + S2B("0h008000FF") + ;
			S2B("0h") + BINTOCX(StyleRecord, "2RS") + S2B("0h0400") + S2B("0h148005FF")

		RETURN lcBiffStyle
	ENDFUNC


	FUNCTION WriteDefaultRowHeight
		* #DEFINE DefRowHeight          0x0225
		LOCAL lqReturn
		lqReturn = S2B("0h")
		IF NOT EMPTY(THIS.nDefaultRowHeight)
			lqReturn = S2B("0h") + ;
				BINTOCX(DefRowHeight, "2RS") + ;
				BINTOCX(4, "2RS") + ;
				BINTOCX(1, "2RS") + ;
				BINTOCX(THIS.nDefaultRowHeight * 20, "2RS")
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
		lqReturn = S2B("0h")
		IF NOT EMPTY(THIS.nDefaultColumnWidth)
			lqReturn = S2B("0h") + ;
				BINTOCX(DefColWidth, "2RS") + ;
				BINTOCX(2, "2RS") + ;
				BINTOCX(THIS.nDefaultColumnWidth, "2RS")
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
		lnCols = ALEN(THIS.aColWidths, 1)
		IF VARTYPE(THIS.aColWidths(1)) <> "N"
			lnNextCol = 1
		ELSE
			lnNextCol = lnCols + 1
		ENDIF
		DIMENSION THIS.aColWidths(lnNextCol, 2)
		THIS.aColWidths(lnNextCol, 1) = tnColumn
		THIS.aColWidths(lnNextCol, 2) = tnWidth
	ENDFUNC


	FUNCTION WriteColumnInfoRecord
		LPARAMETERS tnCol, tnWidth
		* #DEFINE ColumnInfoRecord     0x007D

		tnCol = MAX(0, tnCol - 1)

		LOCAL lcReturn
		lcReturn = S2B("0h") + ;
			BINTOCX(ColumnInfoRecord, "2RS") + ;
			BINTOCX(12, "2RS") + ;
			BINTOCX(tnCol, "2RS") + ;
			BINTOCX(tnCol, "2RS") + ;
			BINTOCX((tnWidth * 256 / 7), "2RS") + ;
			BINTOCX(15, "2RS") + ;
			BINTOCX(0, "2RS") + ;
			BINTOCX(0, "2RS")
		RETURN lcReturn
	ENDFUNC


	FUNCTION WriteCell(tnIndex, tnRow, tnCol, tuValue)
		LOCAL lcType
		lcType = VARTYPE(tuValue)

		DO CASE
			CASE lcType = "C"
				THIS.WriteString(tnIndex, tnRow, tnCol, tuValue)

			CASE lcType = "N"
				THIS.WriteDoble(tnIndex, tnRow, tnCol, tuValue)

			CASE lcType = "D"
				THIS.WriteDate(tnIndex, tnRow, tnCol, tuValue)

			OTHERWISE
				THIS.WriteEmpty(tnIndex, tnRow, tnCol)

		ENDCASE
		RETURN



	FUNCTION WriteString(tnIndex, tnRow, tnCol, tcString)
		IF LEN(tcString) > 255
			tcString = LEFT(tcString, 255)
		ENDIF

		LOCAL lnFXIndex
		lnFXIndex = EVL(THIS.aCells(tnIndex, 10), 0)

		* #DEFINE LabelRecord          0x0204
		THIS.cXLS = THIS.cXLS + ;
			BINTOCX(LabelRecord, "2RS") + BINTOCX(8 + LEN(tcString), "2RS") + ;
			BINTOCX(tnRow - 1, "2RS") + BINTOCX(tnCol - 1, "2RS") + BINTOCX(lnFXIndex, "2RS") + BINTOCX(LEN(tcString), "2RS") + tcString
		RETURN
		RETURN


	FUNCTION WriteInt(tnIndex, tnRow, tnCol, tnValue)
		LOCAL lnValue
		lnValue = BITLSHIFT(tnValue, 2) + 2
		LOCAL lnFXIndex
		lnFXIndex = THIS.aCells(tnIndex, 10)
		THIS.cXLS = THIS.cXLS + ;
			BINTOCX(0x27e, "2RS") + BINTOCX(10, "2RS") + ;
			BINTOCX(tnRow - 1, "2RS") + BINTOCX(tnCol - 1, "2RS") + BINTOCX(lnFXIndex, "2RS") + BINTOCX(lnValue, "4RS")
		RETURN


	FUNCTION WriteDoble(tnIndex, tnRow, tnCol, tnValue)
		* #DEFINE NumberRecord         0x0203
		LOCAL lnFXIndex
		lnFXIndex = THIS.aCells(tnIndex, 10)
		THIS.cXLS = THIS.cXLS + ;
			BINTOCX(NumberRecord, "2RS") + BINTOCX(14, "2RS")     + ; && Headers
			BINTOCX(tnRow - 1, "2RS") + BINTOCX(tnCol - 1, "2RS") + ; && Row, Col
			BINTOCX(lnFXIndex, "2RS")                           + ; && (ushort)cell.FXIndex
			BINTOCX(tnValue, "8SB")
		RETURN


		#IF VERSION(5) > 600
	FUNCTION WriteDate(tnIndex, tnRow, tnCol, tdDate AS DATE)
	#ELSE
	FUNCTION WriteDate(tnIndex, tnRow, tnCol, tdDate)
		tdDate = IIF(VARTYPE(tdDate) = "T", TTOD(tdDate), tdDate)
	#ENDIF
	* #DEFINE NumberRecord         0x0203
	FUNCTION WriteDate(tnIndex, tnRow, tnCol, tdDate AS DATE)
	#ELSE
	FUNCTION WriteDate(tnIndex, tnRow, tnCol, tdDate)
		tdDate = IIF(VARTYPE(tdDate) = "T", TTOD(tdDate), tdDate)
	#ENDIF
	* #DEFINE NumberRecord         0x0203
	LOCAL lnDateValue, lcDateFormat, lnFXIndex, lcReturn
	lnDateValue	= tdDate - {^1900-01-01} + 2
	lnFXIndex	= THIS.aCells(tnIndex, 10)
	lcReturn = S2B("0h") + ;
		BINTOCX(NumberRecord, "2RS") + BINTOCX(14, "2RS") + BINTOCX(tnRow - 1, "2RS") + BINTOCX(tnCol - 1, "2RS") + ;
		BINTOCX(lnFXIndex, "2RS") + ; &&(ushort)cell.FXIndex }
		BINTOCX(lnDateValue, "8SB")
	THIS.cXLS = THIS.cXLS + ;
		lcReturn
	RETURN



	FUNCTION WriteEmpty(tnIndex, tnRow, tnCol)
		LOCAL lnFXIndex
		lnFXIndex = THIS.aCells(tnIndex, 10)
		THIS.cXLS = THIS.cXLS + ;
			BINTOCX(0x0201, "2RS") + BINTOCX(6, "2RS")            + ;  && Header
			BINTOCX(tnRow - 1, "2RS") + BINTOCX(tnCol - 1, "2RS") + ; && Row, Col
			BINTOCX(lnFXIndex, "2RS")
		RETURN

ENDDEFINE



* Support methods for VFP6 compatibility
* By Victor Espina
*
* May 2014
* 
#IF VERSION(5) = 600

	* TRY-CATCH Alternative
	*
	* Instead of:
	*
	* TRY
	*  commands
	* CATCH TO ex
	* ENDTRY
	*
	* We do:
	*
	* TRY()
	*  commands
	* ex = CATCH()
	* ENDTRY()
	*
PROCEDURE TRY
	IF VARTYPE(gcTRYOnError) = "U"
		PUBLIC gcTRYOnError, goTRYEx
	ENDIF
	gcTRYOnError = ON("ERROR")
	goTRYEx		 = NULL
	ON ERROR tryCatch(ERROR(), MESSAGE(), MESSAGE(1), PROGRAM(), LINENO())
ENDPROC
PROCEDURE CATCH(poEx)
	IF PCOUNT() = 1
		poEx = goTRYEx
	ENDIF
	RETURN !ISNULL(goTRYEx)
ENDPROC
PROCEDURE ENDTRY
	IF !EMPTY(gcTRYOnError)
		ON ERROR &gcTRYOnError
	ELSE
		ON ERROR
	ENDIF
ENDPROC
PROCEDURE tryCatch(pnErrorNo, pcMessage, pcSource, pcProcedure, pnLineNo)
	*:Global goTRYEx
	goTRYEx = CREATE("Custom")
	goTRYEx.ADDPROPERTY("errorNo", pnErrorNo)
	goTRYEx.ADDPROPERTY("Message", pcMessage)
	goTRYEx.ADDPROPERTY("Source", pcSource)
	goTRYEx.ADDPROPERTY("Procedure", pcProcedure)
	goTRYEx.ADDPROPERTY("lineNo", pnLineNo)
ENDPROC



* EVL Function 
FUNCTION EVL(puExpr, puDefault)
	RETURN IIF(EMPTY(puExpr), puDefault, puExpr)
ENDFUNC


* GETWORDNUM function
FUNCTION GETWORDNUM(pcList, pnElement, pcSeparator)
	* A previous implementation using ALINES didn't work with hex values
	LOCAL nCount, cElement
	pcSeparator	= EVL(pcSeparator, ",")
	nCount		= OCCURS(pcSeparator, pcList) + 1
	cElement	= ""
	DO CASE
		CASE pnElement = 1
			cElement = LEFT(pcList, AT(pcSeparator, pcList) - 1)

		CASE pnElement = nCount
			cElement = SUBS(pcList, RAT(pcSeparator, pcList) + LEN(pcSeparator))

		CASE pnElement > 1 AND pnElement < nCount
			cElement = SUBS(pcList, AT(pcSeparator, pcList, pnElement - 1) + LEN(pcSeparator))
			cElement = LEFT(cElement, AT(pcSeparator, cElement) - 1)
	ENDCASE
	RETURN cElement
ENDFUNC

#ENDIF


* S2B
* String to binary 
*
* Instead of:
* var = 0h0014
* 
* We use:
* var = S2B("0h0014")
*
PROCEDURE S2B(pcValue)
	IF LEN(pcValue) < 3
		RETURN ""
	ENDIF
	pcValue = SUBSTR(pcValue, 3)
	LOCAL i, cBinValue
	cBinValue = ""
	FOR i = 1 TO LEN(pcValue) STEP 2
		cBinValue = cBinValue + CHR( EVALUATE("0x" + SUBSTR(pcValue, i, 2)) )
	ENDFOR
	RETURN cBinValue
ENDPROC


* BINTOCX
* Substitue for the following BINTOC calls:
* BINTOCX("2RS")
* BINTOCX("4RS")
* BINTOCX("8SB")
*
FUNCTION BINTOCX(pnValue, pcFlags)
	LOCAL cResult
	cResult = ""
	DO CASE
		CASE pcFlags = "4RS"
			cResult = num2dword(pnValue)

		CASE pcFlags = "2RS"
			cResult = LEFT(num2dword(pnValue), 2)

		CASE pcFlags = "8SB"
			cResult = myBinToC8SB(pnValue)
	ENDCASE
	RETURN cResult
ENDFUNC




* num2dword
* Converts a numeric value into a DWORD binary string
*
FUNCTION  num2dword (lnValue)
	#DEFINE m0       256
	#DEFINE m1     65536
	#DEFINE m2  16777216
	LOCAL b0, b1, b2, b3
	b3 = INT(lnValue / m2)
	b2 = INT((lnValue - b3 * m2) / m1)
	b1 = INT((lnValue - b3 * m2 - b2 * m1) / m0)
	b0 = MOD(lnValue, m0)
	RETURN CHR(b0) + CHR(b1) + CHR(b2) + CHR(b3)
ENDFUNC

* Credits for Rick Hodgin
FUNCTION myBinToC8SB
	LPARAMETERS tfValue
	LOCAL lcResult
	lcResult = SPACE(8)
	DECLARE memcpy IN msvcr71.DLL STRING@ DEST, DOUBLE@ SOURCE, INTEGER LENGTH
	memcpy(@lcResult, @tfValue, 8)
	RETURN lcResult
ENDFUNC


* SCANX
* Substitute for VFP9 ASCAN function
*
FUNCTION ASCANX(paArray, puExpr, pnStartElement, pnElementsSearched, pnSearchColumn, pnFlags)
	LOCAL nIndex, i, uValue, j, cType
	nIndex = 0
	j	   = 0
	cType  = VARTYPE(puExpr)
	FOR i = pnStartElement TO ALEN(paArray, 1)
		uValue = paArray[i, pnSearchColumn]
		IF (cType <> "C" AND uValue = puExpr) OR (cType = "C" AND UPPER(uValue) = UPPER(puExpr))
			nIndex = i
			EXIT
		ENDIF
		j = j + 1
		IF j = pnElementsSearched
			EXIT
		ENDIF
	ENDFOR
	RETURN nIndex
ENDFUNC
