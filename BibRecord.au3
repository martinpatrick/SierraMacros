#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=images\autoiticon.ico
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Fileversion=0.0.0.24
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.4.4
 Authors:
 Martin Patrick, Academic Resident Librarian, Miami University
		patricm@miamioh.edu OR martinpatrick@outlook.com
 Jason Paul Michel, User Experience Librarian, Miami University
		micheljp@miamioh.edu
 Becky Yoose, Bibliographic Systems Librarian, Miami University
		b.yoose@gmail.com

 Name of script: BibRecord

 Script Function:
	This script parses and processes data from the bib record as part of the receipt
	cataloging process and chooses whether the receipt team should send to copycat
	or to finalize the item and send it to the shelf.

 Programs used: III Sierra 1.2.1_7, Windows 8.1

 Last revised: 06/2015

 PLEASE NOTE - This script uses a custom UDF library called TSCustomFunction.
			   The script will not run properly (if launched from .au3 file)
			   if that file is not included in the include folder in the
			   AutoIt directory.

The MIT License (MIT)

Copyright (c) 2015 Miami University Libraries

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without
limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------

;######### INCLUDES AND OPTIONS #########
#Include <TSCustomFunction.au3>
AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("SendKeyDelay", 50)
TraySetIcon(@WorkingDir & "\Images\blackB.ico")

;######### DECLARE VARIABLES #########
Dim $REVIEW, $T85, $T86, $T87, $T88, $T89
dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $row1_MASTER, $acqtypecheck, $row2_MASTER, $row3_MASTER
Dim $BCODE3
Dim $CAT_DATE
Dim $LANG
Dim $BCODE1
Dim $COUNTRY
Dim $ccat
Dim $SKIP
Dim $BCODE2
Dim $BIB_LOC_1
Dim $CAT_DATE_R, $TODAY_DATE
Dim $decide
Dim $MARC_LEAD_A, $MARC_LEAD, $MARC_LEAD_L
Dim $040_A, $040, $040_DLC, $040_DNLM, $040_DLCRDA
Dim $042_A, $042, $042_PCC
Dim $050_A, $050_A_N, $050, $050_Z, $050_Z_A, $050_Z_N, $050_Z_L, $050_PZ
Dim $245_2_A, $245_2
Dim $300_A, $300, $300_C_A, $300_C_A_1,$300_C, $FOLIO, $300_E, $300_E_A, $300_E_1, $KNGFO
Dim $050_1, $050_2, $050_2_S, $050_2_K, $050_345
Dim $REF
Dim $IS_IMC
Dim $260_A
Dim $008_D_A, $008_D
Dim $050_D
Dim $260, $260_C
Dim $490_A, $830_A, $440_A
Dim $SF_NAME, $FUND, $multi, $multi_A
dim $BIB_LOC_L2, $BIB_LOC_L1
Dim $300_C_X, $260_C1, $260_L
Dim $LOCATION
Dim $FUND
Dim $BIB_LOC
Dim $SF_590, $SF_7XX, $SF_NAME
Dim $LOCAL
Dim $008_C
Dim $dean
Dim $FUND_A
Dim $multi_loc
Dim $multi_loc_a
Dim $336, $336_A, $336_check, $336_rda_check
Dim $337, $337_A, $337_check, $337_rda_check
Dim $338, $338_A, $338_check, $338_rda_check
Dim $rda_check
Dim $gmd
Dim $counter
Dim $SCFO, $SCOV, $SCFO
Dim $ICODE1_1
Dim $nondlc

;################################ MAIN ROUTINE #################################
;start loading variables from order record script
$LOCATION = _LoadVar("$LOCATION")
$LOCAL = _LoadVar("$LOCAL")
$FUND = _LoadVar("$FUND")
$REF = _LoadVar("$REF")
$T85 = _LoadVar("$T85")
$T86 = _LoadVar("$T86")
$T87 = _LoadVar("$T87")
$T88 = _LoadVar("$T88")
$dean = _LoadVar("$dean")
$multi_loc = _LoadVar("$multi_loc")
$ICODE1_1 = _LoadVar("$ICODE1_1")

;MsgBox(0, "variable check", $ICODE1_1)

$REVIEW = 1
$T89 = ""

;focus on order record screen, go to bib record
;If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;	WinWaitActive ("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Millennium record", "Please open the order record.")
;	Exit
;EndIf

If WinExists("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the order record.")
	Exit
EndIf

;focus on order record screen, go to bib record / Sierra workaround
;If WinExists("[TITLE:Sierra; CLASS:SunAwtFrame]") Then
;	WinActivate("[TITLE:Sierra; CLASS:SunAwtFrame]")
;	WinWaitActive ("[TITLE:Sierra; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Sierra record", "Please open the order record.")
;	Exit
;EndIf

Sleep(0100)
_SendEx("^+e")
;Sleep(0100)
;_SendEx("v")
WinWaitActive ("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0500)

;select all and copy bib record information, parse data into array
_DataCopy()
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;~ ##### start fixed data setting #####
;row 1 includes lang/cat date/ bcode3
$row1_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 1)
$BCODE3 = _ArrayPop($row1_MASTER)
$CAT_DATE = _ArrayPop($row1_MASTER)
$LANG = _ArrayPop($row1_MASTER)
$LANG = StringTrimLeft($LANG, 5)
$CAT_DATE = StringTrimLeft($CAT_DATE, 8)
$BCODE3 = StringTrimLeft($BCODE3, 6)

;row 2 includes skip, bcode1, country
$row2_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 2)
$COUNTRY = _ArrayPop($row2_MASTER)
$BCODE1 = _ArrayPop($row2_MASTER)
$SKIP = _ArrayPop($row2_MASTER)
$SKIP = StringTrimLeft($SKIP, 5)
$BCODE1 = StringTrimLeft($BCODE1, 6)
$COUNTRY = StringTrimLeft($COUNTRY, 7)

;row 3 includes bib location and bcode2
$row3_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 3)
$BCODE2 = _ArrayPop($row3_MASTER)
$BIB_LOC_1= _ArrayPop($row3_MASTER)
$BCODE2 = StringTrimLeft($BCODE2, 6)
$BIB_LOC_1 = StringTrimLeft($BIB_LOC_1, 9)
$BIB_LOC_1 = StringLeft($BIB_LOC_1, 5)
$BIB_LOC_1 = StringStripWS($BIB_LOC_1, 8)
_StoreVar ("$BIB_LOC_1")
;~ ##### end fixed data setting #####

;MsgBox(0, "location", $LOCATION)

;~ ##### start CAT_DATE check #####
$CAT_DATE_R = StringInStr($CAT_DATE, "-  -")
If $CAT_DATE_R = 0 Then
	$decide = MsgBox(4, "Already received", "This item was already received on " & $CAT_DATE & ". Do you want to override the CDATE?")
	Switch $decide
		Case 6
			$TODAY_DATE = "t"
		Case 7
			$TODAY_DATE = "r"
	EndSwitch
EndIf
;~ ##### end CAT_DATE check #####


;~ ##### start MARC LEADER 8 check #####
$MARC_LEAD = _arrayItemString($BIB_ARRAY_MASTER, "MARC Leader	")
$MARC_LEAD =  StringTrimLeft($MARC_LEAD, 12)
$MARC_LEAD_L = StringMid($MARC_LEAD, 18, 1)
If $MARC_LEAD_L = 8 Then
	$T85 = $T85 & "This record has an 8 in the 30th position of the MARC LEADER." & @CR
	$T89 = "SEND TO COPY CATALOGING"
	$REVIEW = 0
EndIf
;~ ##### end MARC LEADER 8 check #####

;~ ##### start DLC/DLC check #####
$040_A = _ArraySearch($BIB_ARRAY_MASTER, "MARC	040	 	 	", 0, 0, 0, 1)
If $040_A > -1 Then
	$040 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $040_A, $040_A)
	$040 =  StringTrimLeft($040, 13)
	$040_DLC = StringInStr($040, "DLC|cDLC")
	$040_DNLM = StringInStr($040, "DNLM/DLC|cDLC")
	$040_DLCRDA = StringInStr($040, "DLC|erda|beng|cDLC")
	$042_A = _ArraySearch($BIB_ARRAY_MASTER, "MARC	042	 	 	", 0, 0, 0, 1)
		If $042_A > -1 Then
			$042 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $042_A, $042_A)
			$042 =  StringTrimLeft($042, 13)
			$042_PCC = StringInStr($042, "pcc")
			If $040_DLC = 0 And $040_DNLM = 0 And $040_DLCRDA = 0 Then
				If $042_PCC = 0 Then
					$T85 = $T85 & "This is a non-DLC/PCC record" & @CR
					$T89 = "SEND TO COPY CATALOGING"
					$REVIEW = 0
				EndIf
			EndIf
		Else
			If $040_DLC = 0 AND $040_DNLM = 0 And $040_DLCRDA = 0 Then
				$T85 = $T85 & "This is a non-DLC record" & @CR
				$T89 = "SEND TO COPY CATALOGING"
				$REVIEW = 0
			EndIf
		EndIf
Else
		$T85 =  $T85 & "This record does not have an 040 field" & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
EndIf
;~ ##### end DLC/DLC check #####

;~ ##### start 050 check #####
$050_A = _ArrayFindAll($BIB_ARRAY_MASTER, "CALL #	050", 0, 0, 0, 1)
$050_A_N = UBound($050_A)
If $050_A_N > 0 Then
	If $050_A_N > 1 Then
		$T85 = $T85 & "This record has multiple 050 fields" & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
	Else
		$050_A = _ArraySearch($BIB_ARRAY_MASTER, "CALL #	050", 0, 0, 0, 1)
		If $050_A > -1 Then
			$050 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $050_A, $050_A)
			$050 =  StringTrimLeft($050, 15)
			$050_Z = StringInStr($050, "Z", 1, 1, 1, 1)
			If $050_Z <> 0 Then
				$050_Z_A = StringSplit($050, "|")
				$050_Z = _ArrayToString($050_Z_A, @TAB, 1, 1)
				$050_Z_N = StringMid($050_Z, 2, 4)
				If $050_Z_N > 1200 Then
					$T85 = $T85 & "The call number is larger than Z1200.  " & @CR
					$T89 = "SEND TO COPY CATALOGING"
					$REVIEW = 0
				EndIf
			EndIf
			$050_PZ = StringInStr($050, "PZ", 1, 1, 1, 1)
			If $050_PZ <>0 AND $LOCATION = "kngli" Then
				$T85 = $T85 & "This record has a PZ in the 050 field in the bibl record, but KNGLI in the order record." & @CR
				$T89 = "SEND to Librarian"
				$REVIEW = 0
			EndIf
		EndIf
	EndIf
Else
	$T85 = $T85 & "This record does not have an 050 field" & @CR
	$T89 = "SEND TO COPY CATALOGING"
	$REVIEW = 0
EndIf
;~ ##### end 050 check #####

;~ ##### start SKIP check #####
$245_2 = _arrayItemString($BIB_ARRAY_MASTER, "TITLE	245")
$245_2 =  StringMid($245_2, 13, 1)
If $SKIP <> $245_2 Then
	MsgBox(64, "Skip does not match 245 second indicator", "Skip does not match the 245 second indicator." & @CR & "Please adjust and click ok to continue.")
EndIf
;~ ##### end SKIP check #####


;~ ##### start 300 check #####
$300_A = _ArraySearch($BIB_ARRAY_MASTER, "DESCRIPT.	300", 0, 0, 0, 1)
If $300_A > -1 Then
	$300 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $300_A, $300_A)
	$300 =  StringTrimLeft($300, 18)
	If $300 = "p.|ccm" Then
		$T85 = $T85 & "This record is a CIP." & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
	Else
		$300_C_A = StringSplit($300, "|")
		$300_C_A_1 = _ArraySearch($300_C_A, "cm", 0, 0, 0, 1)
		$300_C = _ArrayToString($300_C_A, @TAB, $300_C_A_1, $300_C_A_1)
		$300_C_X = StringInStr($300_C, "x")
		If $300_C_X > 0 Then
			$300_C_X = $300_C_X + 2
			$300_C_X = StringMid($300_C, $300_C_X, 2)
		EndIf
		$300_C = StringMid($300_C, 2, 2)
		If $BIB_LOC_1 = "scl" AND $300_C >= 27 Then
			$SCOV = 1
			If $SCOV =1 AND $300_C >= 31 Then
				$SCFO = 1
				$SCOV = 0
			EndIf

			If $BIB_LOC_1 = "kngl" AND $300_C >= 30 Then
				$KNGFO = 1
			EndIf
		ElseIf $300_C = "cm" Then
			$T85 = $T85 & "This record does not have a measurement in 300 |c." & @CR
			$T89 = "SEND TO COPY CATALOGING"
			$REVIEW = 0
		EndIf
		$300_E = StringInStr($300, "|e")
		If $300_E > 0 Then
			If $BIB_LOC = "mul" Then
				$T85 = $T85 & "This Music item has an accompanying material." & @CR
				$T89 = "SEND TO Music Librarian"
				$REVIEW = 0
				$300_E = ""
			Else
			$300_E_A = $300_C_A_1 + 1
			$300_E = _ArrayToString($300_C_A, @TAB, $300_E_A, $300_E_A)
			$300_E = StringTrimLeft($300_E, 1)

			$300_E_1 = StringInStr($300_E, "(")
			$300_E = StringLeft($300_E, $300_E_1)
			$300_E = StringTrimRight($300_E, 1)
			MsgBox(0, "300 check", $300_E)
			EndIf
		Else
			$300_E = 0
		EndIf
	EndIf
Else
	$T85 = $T85 & "This record does not have a 300 field" & @CR
	$T89 = "SEND TO COPY CATALOGING"
	$300_E = 0
	$REVIEW = 0
EndIf
;~ ##### end 300 check #####

;~ ##### start 336 check #####
$336 = _ArraySearch($BIB_ARRAY_MASTER, "NOTE	336", 0, 0, 0, 1)
If $336 > -1 Then
	$336 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $336, $336)
	$336_A = StringSplit($336,"|")
	; _ArrayDisplay($336_A)
	$336_check = _ArraySearch($336_A, "NOTE	336	 	 	text")
	If $336_check = 1 Then
		$336_rda_check = 0
	Else
		$336_rda_check = 1
	EndIf
	; #### MsgBox(1, "$rda check", $336_rda_check)
EndIf

;~ ##### start 337 check #####
$337 = _ArraySearch($BIB_ARRAY_MASTER, "NOTE	337", 0, 0, 0, 1)
If $337 > -1 Then
	$337 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $337, $337)
	$337_A = StringSplit($337,"|")
	; _ArrayDisplay($337_A)
	$337_check = _ArraySearch($337_A, "NOTE	337	 	 	unmediated")
	If $337_check = 1 Then
		$337_rda_check = 0
	Else
		$337_rda_check = 1
	EndIf
	; #### MsgBox(1, "$rda check", $337_rda_check)
EndIf

;~ ##### start 338 check #####
$338 = _ArraySearch($BIB_ARRAY_MASTER, "NOTE	338", 0, 0, 0, 1)
If $338 > -1 Then
	$338 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $338, $338)
	$338_A = StringSplit($338,"|")
	; _ArrayDisplay($338_A)
	$338_check = _ArraySearch($338_A, "NOTE	338	 	 	volume")
	If $338_check = 1 Then
		$338_rda_check = 0
	Else
		$338_rda_check = 1
	EndIf
	; #### MsgBox(1, "$rda check", $338_rda_check)
EndIf

$rda_check = ($338_rda_check + $337_rda_check + $336_rda_check)
; #### MsgBox(1, "$rda check", $rda_check)

if ($rda_check > 0) Then
	#cs
	$gmd = InputBox("GMD Select", "Choose the correct number for this item's GMD and hit enter.  Then put the cursor where the GMD should go and hit F12." & @CR & "0 - GMD has typo; ignore" & @CR & "1 - activity card" & @CR & "2 - art original" & @CR & "3 - art reproduction" & @CR & "4 - braille" & @CR & "5 - cartographic material" & @CR & "6 - chart" & @CR & "7 - diorama" & @CR & "8 - electronic resource" & @CR & "9 - filmstrip" & @CR & "10 - flash card" & @CR & "11 - game" & @CR & "12 - kit" & @CR & "13 - manuscript" & @CR & "14 - microform" & @CR & "15 - microscope slide" & @CR & "16 - model" & @CR & "17 - motion picture" & @CR & "18 - music" & @CR & "19 - picture" & @CR & "20 - realia" & @CR & "21 - slide" & @CR & "22 - sound recording" & @CR & "23 - technical drawing" & @CR & "24 - text" & @CR & "25 - toy" & @CR & "26 - transparency" & @CR & "27 - videorecording", "","","","700")
	Switch $gmd
	    Case "0"
			$gmd = "0"
		Case "1"
			$gmd = "[activity card]"
		Case "2"
			$gmd = "[art original]"
		Case "3"
			$gmd = "[art reproduction]"
		Case "4"
			$gmd = "[braille]"
		Case "5"
			$gmd = "[cartographic material]"
		Case "6"
			$gmd = "[chart]"
		Case "7"
			$gmd = "[diorama]"
		Case "8"
			$gmd = "[electronic resource]"
		Case "9"
			$gmd = "[film strip]"
		Case "10"
			$gmd = "[flash card]"
		Case "11"
			$gmd = "[game]"
		Case "12"
			$gmd = "[kit]"
		Case "13"
			$gmd = "[manuscript]"
		Case "14"
			$gmd = "[microform]"
		Case "15"
			$gmd = "[microscope slide]"
		Case "16"
			$gmd = "[model]"
		Case "17"
			$gmd = "[motion picture]"
		Case "18"
			$gmd = "[music]"
		Case "19"
			$gmd = "[picture]"
		Case "20"
			$gmd = "[realia]"
		Case "21"
			$gmd = "[slide]"
		Case "22"
			$gmd = "[sound recording]"
		Case "23"
			$gmd = "[technical drawing]"
		Case "24"
			$gmd = "[text]"
		Case "25"
			$gmd = "[toy]"
		Case "26"
			$gmd = "[transparency]"
		Case "27"
			$gmd = "[videorecording]"
		 EndSwitch
   If $gmd <> "0" Then
;	MsgBox(1,"$gmd check", $gmd)
	Local $hDLL = DllOpen("user32.dll")
	$counter = 0
	While ($counter = 0)
		If _IsPressed("7B", $hDLL) Then
		While _IsPressed("7B", $hDLL)
			Sleep(050)
        WEnd
        _SendEx("{INS}" & "|h" & $gmd )
		$counter = 1
		ExitLoop
		EndIf
	WEnd
DllClose($hDLL)
EndIf
#ce
$T85 = $T85 & "This record does has issues in the 33X Fields." & @CR
	$T89 = "SEND TO COPY CATALOGING"
EndIf

;~ ##### start acq call number/location assignment and check #####
If $BIB_LOC_1 = "acq" Then
	$050_1 = StringLeft($050, 1)
	$050_2 = StringMid($050, 2, 1)
	Switch $050_1
		Case "M"
			$LOCATION = "muli"
		Case "N"
			$LOCATION = "aali"
		Case "Q", "R", "S"
			$LOCATION = "scli"
		Case "T"
			If $050_2 = "R" Then
				$LOCATION = "aali"
			Else
			$LOCATION = "scli"
			EndIf
		Case "G"
			$050_2_S = StringRegExp($050_2, "[A-CE4-9]")
			$050_2_K = StringRegExp($050_2, "[FGHKLMNRTV12]")
			If $050_2_S > 0 Then
				$LOCATION = "scli"
			ElseIf $050_2_K > 0 Then
				$LOCATION = "kngli"
			ElseIf $050_2 = 3 Then
				$050_345 = StringMid($050, 3, 3)
				If $050_345 >= 160 Then
					$LOCATION = "scli"
				Else
					$LOCATION = "kngli"
				EndIf
			EndIf
		Case "P"
			If $050_2 = "Z" Then
				$T85 = $T85 & "This record has a PZ in the 050 field in the bibl record, but KNGLI in the order record." & @CR
				$T89 = "SEND to Catalog Librarian"
				$REVIEW = 0
			Else
				$LOCATION = "kngli"
			EndIf
		Case Else
			$LOCATION = "kngli"
	EndSwitch
	_StoreVar("$LOCATION")
EndIf

Sleep(1500)

If $REF = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "r")
	_StoreVar("$LOCATION")
EndIf
;~ ##### end check #####

;~ ##### start multi check #####
If $BIB_LOC_1 = "multi" Then
	$T85 = $T85 & "This record has a location of multi" & @CR
	$T89 = "SEND TO COPY CATALOGING OR ADDED COPIES"
	$REVIEW = 0
EndIf
;~ ##### end multi check #####

;~ ##### start imju check #####
If $LOCATION = "imju" Then
	$decide = InputBox("Choose IMC Juv location", "Choose which type of juv book this item is. Type in the corresponding letter." & @CR & "A: Juv (Fic and Nonfic)" & @CR & "B: Easy" & @CR & "C: Big Book"  & @CR & "D: IMC Reader" & @CR & "E: IMC LangMat", "", "", 200, 210)
	$decide = StringUpper($decide)
	Switch $decide
		Case "A"
			$LOCATION = "imju"
		Case "B"
			$LOCATION = "imje"
		Case "C"
			$LOCATION = "imjb"
		Case "D"
			$LOCATION = "imrdr"
		Case "E"
			$LOCATION = "imlng"
	EndSwitch
			_StoreVar("$LOCATION")
EndIf
;~ ##### end imju check #####

;~ ##### start folio check #####
#cs
$IS_IMC = StringInStr($LOCATION, "im")
If $FOLIO = 1 AND $IS_IMC = 0 And $LOCATION <> "scli" AND $KNGFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
	_StoreVar("$LOCATION")
 ElseIf $FOLIO = 1 And $IS_IMC = 0 And $LOCATION = "scli" Then
	$LOCATION = StringReplace($LOCATION, "li", "ov")
	_StoreVar("$LOCATION")
ElseIf $FOLIO = 1 And $IS_IMC = 1 And $SCFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
	_StoreVar("LOCATION")
EndIf
;~ ##### end folio check #####
#ce
If $SCOV = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "ov")
EndIf
If $SCFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf
If $KNGFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf

;~ ##### start dean material location check #####
;this sets the default bib location to kngli for dean's office materials
If $dean = 1 Then
	$LOCATION = "kngli"
EndIf
;~ ##### end dean material location check #####

;##### RUN BIB LOCATION SCRIPT #####
Run(@WorkingDir & "ILOC.exe")
	While ProcessExists("ILOC.exe")
		Sleep(0400)
	WEnd
$BIB_LOC = _LoadVar("$BIB_LOC")
$REF = _LoadVar("$REF")
Sleep(0300)
;##### END RUN BIB LOCATION SCRIPT #####

;~ ##### start bcode1 check #####
$BCODE1 = StringLeft($BCODE1, 1)
If $BCODE1 = "s" OR $BCODE1 = "b" Then
	$T85 = $T85 & "This is a SERIAL" & @CR
	$T89 = "SEND TO SERIALS"
	$REVIEW = 0
EndIf
;~ ##### end bcode1 check #####

;~ ##### start bcode2 check #####
$BCODE2 = StringLeft($BCODE2, 1)
If $BCODE2 = "c" Then
	$T85 = $T85 & "This record is for a musical score" & @CR
	$T89 = "SEND to Music Cataloging"
	$REVIEW = 0
EndIf
;~ ##### end bcode2 check #####

;~ ##### start bcode3 check #####
;NOT LOOPED YET
If $LOCAL = -1 Then
	If $BCODE3 <> "-" Then
		MsgBox(64, "BCODE3 is not set to '-'", "BCODE3 is not set to '-'." & @CR & "Please correct and click ok to continue.")
	EndIf
ElseIf $dean = 1 Then
	If $BCODE3 <> "s" Then
		MsgBox(64, "BCODE3 is not set to 's'", "BCODE3 is not set to 's'." & @CR & "Please correct and click ok to continue.")
	EndIf
EndIf
;~ ##### end bcode3 check #####

;~ ##### start date check #####
$260_A = _ArraySearch($BIB_ARRAY_MASTER, "IMPRINT	260", 0, 0, 0, 1)
If $MARC_LEAD_L <> 8 AND $050_A_N = 1 AND $260_A > -1 Then
	;search for 008 date
	$008_D_A = _ArraySearch($BIB_ARRAY_MASTER, "MARC	008", 0, 0, 0, 1)
	If $008_D_A > -1 Then
		$008_D = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $008_D_A, $008_D_A)
		$008_D =  StringMid($008_D, 21, 4)
		$008_D = StringStripWS($008_D, 8)
		;check for conference entry
		$008_C = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $008_D_A, $008_D_A)
		$008_C =  StringMid($008_C, 43, 1)
		$008_C = StringStripWS($008_C, 8)
	EndIf
	;search for 050 date
	$050_D = StringRight($050, 4)
	$050_D = StringStripWS($050_D, 8)
	;search for 260 date
	$260 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $260_A, $260_A)
	$260_L = StringLen($260)
	$260_C1 = StringInStr($260, "|c")
	$260_C = StringMid($260, $260_C1)
	$260_C = StringTrimLeft($260_C, 2)
	$260_C = StringRegExpReplace($260_C, "[][c?]", "")
	$260_C = StringStripWS($260_C, 8)
	If $260_C <> $008_D OR $008_D <> $050_D OR $260_C <> $050_D Then
		;conference check
		If $008_C <> 1 Then
			$T85 = $T85 & "The 008, 050, and 260 |c dates do not match." & @CR
			$T89 = "SEND TO COPY CATALOGING"
			$REVIEW = 0
		EndIf
	EndIf
EndIf
;~ ##### end date check #####

;~ 	##### start series check #####
$440_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	440", 0, 0, 0, 1)
$490_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	490", 0, 0, 0, 1)
$830_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	830", 0, 0, 0, 1)
If $REVIEW = 1 Then
	If $490_A > -1 AND $830_A = -1 Then
		$T85 = $T85 & "The 490 field does not have a corresponding 830 field." & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
	EndIf
	If $440_A > -1 OR $830_A > -1 Then
		$decide = MsgBox(4, "Check the series.", "Check the series." & @CR & "Should this item be classed separately?")
		Switch $decide
			Case 7 ; 7 = no
				$T88 = $T88 & "Classed together." & @CR
				$T89 = "SEND TO SERIALS"
				$REVIEW = 0
				$LOCATION = "none"
			Case 6 ; 6 = yes, should be classed separately

		EndSwitch
		;WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
		WinActivate("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	EndIf
EndIf
;~ 	##### end series check #####

;START SAVING VARIABLES FOR ITEM SCRIPTS
_StoreVar("$300_E")
_StoreVar("$LOCATION")
_StoreVar("$BIB_LOC_1")

Dim $slash, $4
;~ ##### START PROCESSING RECORD #####
If $REVIEW = 1 Then
	;~ ##### start Special Fund check #####
	#cs
		Fun happy fund parsing workaround!
		It turns out that the new fund formatting for 2009B played havoc with
		parsing out fund codes for single fund purchases. Instead of having a
		space between the fund code and the fund name, they are lumped together
		in one solid line. I have a workaround that checks for the first digit
		of the fund line to determine if it's a special fund code. This in part
		works because all but one special fund start with "4" and none have
		a slash in them.

		As far as I know, the multi fund parsing has not run into the same
		issue from the 2009B upgrade, so I have left the original parsing
		intact.

	$slash = StringInStr($FUND, "/")
	$4 = StringLeft($FUND, 1)
	If $slash = 0 And $4 = 4 Then
		$FUND_A = StringRegExp($FUND, "(4\D{1,3})", 1)
		$FUND = _ArrayPop($FUND_A)
	EndIf
	$FUND = StringStripWS($FUND, 8)
	_StoreVar("$FUND")
	If $FUND <> "multi" then
		Run(@DesktopDir & "\SierraReceipt - gmd\Special Fund List.exe")
			While ProcessExists("Special Fund List.exe")
				Sleep(0400)
			WEnd
		$SF_NAME = _LoadVar("$SF_NAME")
		If $SF_NAME <> "none" Then
			$SF_590 = _LoadVar("$SF_590")
			$SF_7XX = _LoadVar("$SF_7XX")
			;WinActive("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")
			Sleep(0100)
			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}n")
			Sleep(0100)
			_SendEx("590{TAB}")
			sleep(0100)
			Switch $BIB_LOC
				Case "kngl"
					_SendEx("King copy{SPACE}")
				Case "scl"
					_SendEx("Science copy{SPACE}")
				Case "aal"
					_SendEx("ArtArch copy{SPACE}")
				Case "imc"
					_SendEx("IMC copy{SPACE}")
				Case "mul"
					_SendEx("Music copy{SPACE}")
			EndSwitch
			_SendEx($SF_590)
			Sleep(0300)
			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}b")
			Sleep(0100)
			_SendEx($SF_7XX)
		EndIf
	ElseIf $FUND = "multi" Then
		$multi = _LoadVar("$multi")
		$multi_A = StringSplit($multi, ",")
		Do
			$FUND = _ArrayPop($multi_A)
			$FUND = StringStripWS($FUND, 8)
				Run(@DesktopDir & "\Receipt\Special Fund List.exe")
					While ProcessExists("Special Fund List.exe")
						Sleep(0400)
					WEnd
				$SF_NAME = _LoadVar("$SF_NAME")
			If $SF_NAME <> "none" Then
				$SF_590 = _LoadVar("$SF_590")
				$SF_7XX = _LoadVar("$SF_7XX")
				;$SF_7XX = WinActive("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")
				_SendEx("^{END}")
				Sleep(0100)
				_SendEx("{ENTER}n")
				Sleep(0100)
				_SendEx("590{TAB}")
				sleep(0100)
				Switch $BIB_LOC
					Case "kngl"
						_SendEx("King copy{SPACE}")
					Case "scl"
						_SendEx("Science copy{SPACE}")
					Case "aal"
						_SendEx("ArtArch copy{SPACE}")
					Case "imc"
						_SendEx("IMC copy{SPACE}")
					Case "mul"
						_SendEx("Music copy{SPACE}")
				EndSwitch
				_SendEx($SF_590)
				Sleep(0300)
				_SendEx("^{END}")
				Sleep(0100)
				_SendEx("{ENTER}b")
				Sleep(0100)
				_SendEx($SF_7XX)
			EndIf
		Until $multi_A = ""
	EndIf
	;~ ##### end Special Fund check #####
	#ce
	SLEEP(0400)
	If $BIB_LOC = "multi" Then ;usually if item is for reference
		$BIB_LOC_L1 = _LoadVar("$BIB_LOC_L1")
		$BIB_LOC_L2 = _LoadVar("$BIB_LOC_L2")
		_SendEx("^{HOME}")
		Sleep(0100)
		_SendEx("{DOWN 2}")
		Sleep(0300)
		_SendEx($BIB_LOC)
		Sleep(0100)
		WinWaitActive("[TITLE:Edit Data; CLASS:SunAwtDialog]")
		Sleep(0300)
		_SendEx($BIB_LOC_L1)
		Sleep(0300)
		_SendEx("!a")
		Sleep(0300)
		_SendEx($BIB_LOC_L2)
		Sleep(0300)
		_SendEx("!o")
	ElseIf $multi_loc <> "" Then
		_SendEx("^{HOME}")
		Sleep(0100)
		_SendEx("{DOWN 2}")
		Sleep(0300)
		_SendEx("multi")
		Sleep(0100)
		WinWaitActive("[TITLE:Edit Data; CLASS:SunAwtDialog]")
		Sleep(0300)
		$multi_loc_a = StringSplit($multi_loc, ",")
		Do
			$LOCATION = _ArrayPop($multi_loc_a)
			$LOCATION = StringRegExpReplace($LOCATION, "[(\d)]", "")
			$LOCATION = StringStripWS($LOCATION, 8)
			_StoreVar("$LOCATION")
			Run(@WorkingDir & "ILOC.exe")
			While ProcessExists("ILOC.exe")
				Sleep(0400)
			WEnd
			$BIB_LOC = _LoadVar("$BIB_LOC")
			$REF = _LoadVar("$REF")
			If $BIB_LOC = "multi" Then ; if item is for reference
				$BIB_LOC_L1 = _LoadVar("$BIB_LOC_L1")
				$BIB_LOC_L2 = _LoadVar("$BIB_LOC_L2")
				Sleep(0300)
				_SendEx($BIB_LOC_L1)
				Sleep(0300)
				_SendEx("!a")
				Sleep(0300)
				_SendEx($BIB_LOC_L2)
			Else
				Sleep(0300)
				_SendEx($BIB_LOC)
			EndIf
			Sleep(0300)
			_SendEx("!a")
			Sleep(0300)
		Until $multi_loc_a = ""
		Sleep(0300)
		_SendEx("!o")
	Else
		Sleep(0400)
	EndIf
	MsgBox(64, "Review the content lines before adding item.", "Review the content lines before adding item.", 5) ;done!
	$ccat = 1
	_StoreVar("$ccat")
Else
	If $T85 <> "" OR $T86 <> "" OR $T87 <> "" Or $T88 <> "" Or $T89 <> "" Then
		MsgBox(0, "Receipt Cataloging", $T85 & @CR & $T86 & @CR & $T87 & @CR & $T88 & @CR & $T89)
	EndIf
	Sleep(0100)
	_SendEx("^S")
	Sleep(100)
	$nondlc = 1

	_StoreVar("$nondlc")

	Run("ItemCreate.exe")
EndIf
