#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=images\autoiticon.ico
#AutoIt3Wrapper_UseX64=n
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

 Name of script: FindRecord

 Script Function:
	This script begins the process, and responds based on whether a barcode
	or ISBN is scanned to decide what to do.

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
Opt("WinTitleMatchMode", 1)
Opt("WinSearchChildren", 1)
Opt("MustDeclareVars", 1)
TraySetIcon(@WorkingDir & "\Images\take1new.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $LOCAL
Dim $245, $245_A2
Dim $ISBN
Dim $300, $300_E, $300_E1
Dim $winsize
Dim $decide
dim $barcode
Dim $UPD
Dim $ITEM_REC_PREP, $ITEM_REC, $ITEM_ARRAY_MASTER
Dim $row2_MASTER
Dim $ICODE1
Dim $300
Dim $OCLC_NUM
Dim $VOL
dim $LAB_LOC
Dim $ACCOMP
Dim $dean
Dim $row3_MASTER, $BCODE2, $BIB_LOC_1
Dim $LEADER, $LEADER_3
Dim $TITLE, $245_A2, $245DASH
Dim $SF_NAME, $SF_590, $SF_7XX
Dim $fundrow_MASTER, $ORD_REC, $ORD_REC_PREP, $ORD_ARRAY_MASTER
Dim $FUND, $CLAIM, $STATUS, $FUND2, $RLOC, $EPRICE, $LOCATION, $REF
DIM $COP, $ILOVC, $LPAT, $row1_MASTER
Dim $RDATE, $CODE4, $ACQTYPE, $acqtypecheck, $BLOC, $FORM, $CDATE, $TLOC, $ODATE, $COPIES
Dim $VENDOR, $ORDNOTE, $CODE1, $row4_MASTER, $row5_MASTER, $row6_MASTER, $row7_MASTER, $row8_MASTER
Dim $ORDTYPE, $ORD_LANG, $CODE2, $RACTION, $CODE3, $OLOCATION
Dim $clipcheck
Dim $copycat, $advance

;################################ MAIN ROUTINE #################################
;check to see if an existing record is open, close it and clear buffer

;If WinExists("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [bio][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
;	WinKill("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [bio][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;EndIf
;If WinExists("[REGEXPTITLE:\A[bio][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
;	WinKill ("[REGEXPTITLE:\A[bio][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;EndIf
_ClearBuffer()
Sleep(0300)

; if in multi window mode, make sure to close exising item or bib record
If WinExists("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
WinActivate("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]")

$winsize = WinGetPos("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
_SendEx("^s")
Sleep(0200)
_SendEx("!q")
sleep(0300)
EndIf

;focus on main search screen
If WinExists("[TITLE:Sierra; CLASS:SunAwtFrame]") Then
	WinActivate("[TITLE:Sierra; CLASS:SunAwtFrame]")
;Else
	;MsgBox(64, "nope", "Please log into Sierra.")
	;Exit
EndIf

;ask for ISBN, search III
$ISBN = InputBox("Scan something!", "Scan the item's ISBN or barcode", "")

If $ISBN = "" Then
	Exit
Else
$decide = StringLen($ISBN)
If $decide = 14 Then
	;msgbox(0, "$ISBN", "ISBN Is a Barcode!")
	_IIIsearch("b", $ISBN)
Else
	;msgbox(0, "$ISBN", "ISBN Is not a Barcode!")
_IIIsearch("i", $ISBN)
$decide = 1
	EndIf
	EndIf

If $decide = 1 Then
;go to bib record
;~ WinWaitActive("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
;~ Sleep(0200)
;WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;WinWaitActive("[TITLE:Sierra; CLASS:SunAwtFrame]")
WinWaitActive("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0400)
_SendEx("!g")
Sleep(0400)
_SendEx("!v")

;WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;WinWaitActive("[TITLE:Sierra; CLASS:SunAwtFrame]")
WinWaitActive("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")

;select all and copy bib record information, parse data into array
Sleep(0400)
_DataCopy()
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;245 title search
$245 =  _arrayItemString($BIB_ARRAY_MASTER, "TITLE	245")
$245_A2 = StringSplit($245, "|")
$245 =  _arrayItemString($245_A2, "TITLE	245")
$245 = StringTrimLeft($245, 14)

;300 accompanying material search for internal note
$300 = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300")
$300_E = StringInStr($300, "|e")
If $300_E > 0 Then
	$300_E = StringMid($300, $300_E)
	$300_E = StringTrimLeft($300_E, 2)
	$300_E1 = StringInStr($300_E, "(")
	Switch $300_E1
		case 10
			$300_E = StringLeft($300_E, 10)
			$300_E = StringTrimRight($300_E, 1)
			$300_E = StringStripCR($300_E)
		case 12
			$300_E = StringLeft($300_E, 12)
			$300_E = StringTrimRight($300_E, 1)
			$300_E = StringStripCR($300_E)
		case 13
			$300_E = StringLeft($300_E, 13)
			$300_E = StringTrimRight($300_E, 1)
			$300_E = StringStripCR($300_E)
	EndSwitch
	#cs
	If $300_E1 > 0 Then
		$300_E = StringLeft($300_E, $300_E1)
		$300_E = StringTrimRight($300_E, 1)
	EndIf
	#ce
;Else
	;$300_E = 0
EndIf
_StoreVar("$300_E") ;store for OrderRec use to create internal note

;close record and search III catalog with title
Sleep(0100)
;WinKill("[REGEXPTITLE:\A[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;WinKill("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0100)
_IIIsearch("t", $245)
Sleep(0800)

;dup check - if the order record doesn't pop up in a certain amount of time, then script prompts for a dup check
;If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
IF WinExists("[REGEXPTITLE:[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")	Then
	Exit
Else
	MsgBox(0,"Check possible dup", "There are similar titles in the system. Please select correct title.")
EndIf



;select all and copy bib record information, parse data into array
_DataCopy()
_SendEx("^{HOME}") ;de-highlights record. common courtesy for the catalogers
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "Fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "Fnord", 1)
_StoreVar("$BIB_ARRAY_MASTER")


;row 3 includes bib location and bcode2
$row3_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 3)
$BCODE2 = _ArrayPop($row3_MASTER)
$BIB_LOC_1= _ArrayPop($row3_MASTER)
$BIB_LOC_1 = StringTrimLeft($BIB_LOC_1, 9)
$BIB_LOC_1 = StringLeft($BIB_LOC_1, 5)
$BIB_LOC_1 = StringStripWS($BIB_LOC_1, 8)

;get encoding level from leader, if level 3 open OCLC record
$LEADER = _arrayItemString($BIB_ARRAY_MASTER, "MARC Leader	")
$LEADER =  StringTrimLeft($LEADER, 12)
$LEADER = StringStripWS($LEADER, 8)
$LEADER_3 = StringMid($LEADER, 16, 1)

;300 search - if incomplete open OCLC record
$300 = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300	 	 	")
$300 =  StringTrimLeft($300, 18)
$300 = StringStripWS($300, 8)

;(300 is incomplete OR item is Middletown) AND encoding level != 3 - search OCLC by OCLC #
If ($300 = "p.|ccm" Or $300 = "p.;|ccm" Or $300 = -1  or $300 = "p.cm" OR $ICODE1 = 83) AND $LEADER_3 <> 3 Then
	$OCLC_NUM = _arrayItemString($BIB_ARRAY_MASTER, "OCLC #	001	 	 	")
	$OCLC_NUM = StringTrimLeft($OCLC_NUM, 15)
	_OCLCSearch("{#}", $OCLC_NUM)
EndIf

;Item is STO OR encoding level = 3 - search OCLC by title (subfield a)
If $ICODE1 = 3 OR $LEADER_3 = 3 Then
	$TITLE = _arrayItemString($BIB_ARRAY_MASTER, "TITLE	")
	If $LEADER_3 = 3 Then
		$TITLE = StringTrimLeft($TITLE, 6)
		$245_A2 = StringSplit($TITLE, "|")
		$TITLE = _arrayItemString($245_A2, "245	")
		$TITLE = StringTrimLeft($TITLE, 8)
		$245DASH = StringRegExp($TITLE, "[:/]", 0)
		If $245DASH = 1 Then
			$TITLE = StringTrimRight($TITLE, 2)
		EndIf
	Else
		$TITLE = StringTrimLeft($TITLE, 6)
	EndIf
	_OCLCSearch("ti:", $TITLE)
EndIf

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Else ;if you scanned a barcode instead of an ISBN, this process will run instead
	$ITEM_REC = ""
While $ITEM_REC = "" ;Then ;this prevents the macro from moving on too quickly. basically, it can't move on until it's copied the item
	;If WinExists("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
WinActivate("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
WinWaitActive("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")

$copycat = 1
_StoreVar("$copycat")


$winsize = WinGetPos("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
;EndIf
_DataCopy()
$ITEM_REC = ClipGet()
$ITEM_REC_PREP = StringRegExpReplace ($ITEM_REC, "[\r\n]+", "fnord")
$ITEM_ARRAY_MASTER = StringSplit($ITEM_REC_PREP, "fnord", 1)

;grab item location in case it's from standing orders or doesn't have an order record for some reason
$row1_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 1)
$LOCATION = _ArrayPop($row2_MASTER)
$LPAT = _ArrayPop($row2_MASTER)
$COP = _ArrayPop($row2_MASTER)
$LOCATION = StringTrimLeft($LOCATION, 8)
$LOCATION = StringLeft($LOCATION, 5)
$LOCATION = StringStripWS($LOCATION, 8)

;search for ICODE - determines UPD in 947
$row2_MASTER = _IIIfixed($ITEM_ARRAY_MASTER, 2)
$ICODE1 = _arrayItemString ($row2_MASTER, "CODE1")
$ICODE1 = StringTrimLeft($ICODE1, 6)
$ICODE1 = StringStripWS($ICODE1, 8)
_StoreVar("$ICODE1")

If $ICODE1 = 2 OR $ICODE1 = 83 Then
	$UPD = 0 ;no upd code in 947
Else
	$UPD = 1 ;upd code in 947
EndIf

;determine if additional item records need to be created
$VOL = _ArraySearch($ITEM_ARRAY_MASTER, "VOLUME", 0, 0, 0, 1)
If $VOL > -1 Then
	$VOL = 1
EndIf

;determine if item is dean's office item
$dean = _ArraySearch($ITEM_ARRAY_MASTER, "Catalog for Dean's Office", 0, 0, 0, 1)
If $dean > -1 Then
	$dean = 1
EndIf

#cs -- deprecated as printer no longer needs this field
;determine if label location is present
$LAB_LOC = _ArraySearch($ITEM_ARRAY_MASTER, "LABEL LOC", 0, 0, 0, 1)
If $LAB_LOC > -1 Then
	$LAB_LOC = 1 ;label location not present, will determine later on if record needs one
Else
	$LAB_LOC = 0
EndIf
#ce

;determine if item has accompying item note, will determine later on if record needs one
$ACCOMP = _ArraySearch($ITEM_ARRAY_MASTER, "IN POCKET", 0, 0, 0, 1)
_StoreVar("$ACCOMP")
If $ITEM_REC <> "" Then ExitLoop
Wend



Sleep(0200)



WinWaitActive("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;check order record for fund codes! whee!!! check all the things!!
Sleep(0100)
_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("!l")
Sleep(0600)


WinWaitActive("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")



;select all and copy order record information, parse data into array
_SendEx("^{HOME}")
Sleep(0200)
_SendEx("^a")
Sleep(0100)
_SendEx("^a")
Sleep(0100)
_SendEx("^c")
Sleep(0300)

$ORD_REC = ClipGet()
$ORD_REC_PREP = StringRegExpReplace($ORD_REC, "[\r\n]+", "fnord")
$ORD_ARRAY_MASTER = StringSplit($ORD_REC_PREP, "fnord", 1)

If $ACCOMP = -1 Then

$ACCOMP = _ArraySearch($ORD_ARRAY_MASTER, "Rec'd", 0, 0, 0, 1)

_StoreVar("$ACCOMP")
EndIf


;~ ##### start fixed data setting #####
;row 1 includes acqtype/code4/rdate
$row1_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 1)
$RDATE = _ArrayPop($row1_MASTER)
$CODE4 = _ArrayPop($row1_MASTER)
$ACQTYPE = _ArrayPop($row1_MASTER)
$RDATE = StringTrimLeft($RDATE, 5)
$CODE4 = StringTrimLeft($CODE4, 5)

;~ Bloody acqtype workaround - sometimes the A in Acq type doesn't get captured
$acqtypecheck = StringInStr($ACQTYPE, "ACQ TYPE")
If $acqtypecheck <> 0 Then
	$ACQTYPE = StringTrimLeft($ACQTYPE, 9)
EndIf

;row 2 includes location, eprice, rloc
$row2_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 2)
$RLOC = _ArrayPop($row2_MASTER)
$EPRICE = _ArrayPop($row2_MASTER)
$LOCATION = _ArrayPop($row2_MASTER)
$RLOC = StringTrimLeft($RLOC, 4)
$EPRICE = StringTrimLeft($EPRICE, 7)
$LOCATION = StringTrimLeft($LOCATION, 9)
$LOCATION = StringLeft($LOCATION, 5)
$LOCATION = StringStripWS($LOCATION, 8)
_StoreVar("$LOCATION")
$OLOCATION = $LOCATION
$OLOCATION = StringReplace($OLOCATION, "li", "l")
_StoreVar("$OLOCATION")

;MsgBox(0, "variables", $RLOC & $EPRICE & $LOCATION)

;row 3 includes cdate, form, bloc
$row3_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 3)
$BLOC = _ArrayPop($row3_MASTER)
$FORM = _ArrayPop($row3_MASTER)
$CDATE = _ArrayPop($row3_MASTER)
$BLOC = StringTrimLeft($BLOC, 4)
$FORM = StringTrimLeft($FORM, 4)
$CDATE = StringTrimLeft($CDATE, 6)

;row 4 includes claim, fund, status
$row4_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 4)
$STATUS = _ArrayPop($row4_MASTER)
$FUND = _ArrayPop($row4_MASTER)
$CLAIM = _ArrayPop($row4_MASTER)
$STATUS = StringTrimLeft($STATUS, 6)
$FUND = StringTrimLeft($FUND, 4)
$CLAIM = StringTrimLeft($CLAIM, 6)

$FUND2 = StringLeft($FUND, 3); this is for checking the fund code against the special fund list
_StoreVar("$FUND2")
;msgbox(0, "FUND2", $FUND)

;row 5 includes copies, odate, tloc
$row5_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 5)
$TLOC = _ArrayPop($row5_MASTER)
$ODATE = _ArrayPop($row5_MASTER)
$COPIES = _ArrayPop($row5_MASTER)
$TLOC = StringTrimLeft($TLOC, 4)
$ODATE = StringTrimLeft($ODATE, 5)
$COPIES = StringTrimLeft($COPIES, 7)

;row 6 includes code1, ordnote, vendor
$row6_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 6)
$VENDOR = _ArrayPop($row6_MASTER)
$ORDNOTE = _ArrayPop($row6_MASTER)
$CODE1 = _ArrayPop($row6_MASTER)
$VENDOR = StringTrimLeft($VENDOR, 6)
$ORDNOTE = StringTrimLeft($ORDNOTE, 8)
$CODE1 = StringTrimLeft($CODE1, 6)

;row 7 includes code2, ordtype, lang
$row7_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 7)
$ORD_LANG = _ArrayPop($row7_MASTER)
$ORDTYPE = _ArrayPop($row7_MASTER)
$CODE2 = _ArrayPop($row7_MASTER)
$ORD_LANG = StringTrimLeft($ORD_LANG, 4)
$ORDTYPE = StringTrimLeft($ORDTYPE, 8)
$CODE2 = StringTrimLeft($CODE2, 6)

;row 8 includes code3, raction, volumes
$row8_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 8)
$VOL = _ArrayPop($row8_MASTER)
$RACTION = _ArrayPop($row8_MASTER)
$CODE3 = _ArrayPop($row8_MASTER)
$VOL = StringTrimLeft($VOL, 7)
$RACTION = StringTrimLeft($RACTION, 7)
$CODE3 = StringTrimLeft($CODE3, 6)
;~ ##### end fixed data setting #####



;msgbox(0, "loc", $LOCATION)

If $LOCATION = "kngr" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATION = "aar" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATION = "scr" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATION = "mur" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATIOn = "imr" Then
	$REF = 1
	_StoreVar("$REF")
Else
	$REF = 0
	_StoreVar("$REF")
EndIf



;~ ##### start local use check #####
;used to determine BCODE3 and ITYPE codes
$LOCAL = _ArraySearch($ORD_ARRAY_MASTER, "LOCAL USE ONLY", 0, 0, 0, 1)
;~ ##### end local use check #####

Run("Special Fund List.exe")
While ProcessExists("Special Fund List.exe")
	Sleep(0400)
WEnd
$SF_NAME = _LoadVar("$SF_NAME")
$SF_590 = _LoadVar("$SF_590")
$SF_7XX = _LoadVar("$SF_7XX")

Sleep(0500)

MsgBox(0, "Order notes", "Please make a mental note of any relevant order record notes!")


;going to bib record
Sleep(0100)
_SendEx("^+e")

;wait for bib record focus
;WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")

WinWaitActive("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")

If $ICODE1 = "" Then
	MsgBox(0, "Uh-oh!", "It seems Sierra flaked out on this run. Please close the bib record and start the macro again before proceeding.")
	Exit
EndIf


;store variables for ItemRec script
_StoreVar("$UPD")
_StoreVar("$LAB_LOC")
_StoreVar("$VOL")
_StoreVar("$ICODE1")
_StoreVar("$ACCOMP")
_StoreVar("$BIB_LOC_1")
_StoreVar("$LOCAL")
_StoreVar("$dean")
_StoreVar("$SF_NAME")
_StoreVar("$SF_590")
_StoreVar("$SF_7XX")

If $ICODE1 = "" Then
	MsgBox(0, "Uh-oh!", "It seems Sierra flaked out on this run. Please close the bib record and start the macro again before proceeding.")
	Exit
EndIf
EndIf

Exit
