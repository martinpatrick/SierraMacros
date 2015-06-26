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

 Name of script: OrderRecord

 Script Function:
	This script parses and processes data from the order record display,
	storing data in variables that will be used later to finalize the process.

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
#include "TSCustomFunction.au3"
AutoItSetOption("WinTitleMatchMode", 2)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("SendKeyDelay", 50)
TraySetIcon(@WorkingDir & "\Images\blueO.ico")

;######### DECLARE VARIABLES #########
Dim $winsize
Dim $nondlc
Dim $ORD_REC, $ORD_REC_PREP, $ORD_ARRAY_MASTER
Dim $300_E, $300_E_E
Dim $LOCAL
Dim $ACQTYPE
Dim $CODE4
Dim $RDATE
Dim $OLOCATION
Dim $LOCATION
Dim $EPRICE
Dim $RLOC
Dim $CDATE
Dim $FORM
Dim $BLOC
Dim $CLAIM
Dim $FUND
Dim $STATUS
Dim $CODE1
Dim $ORDNOTE
Dim $ccat
Dim $VENDOR
Dim $CODE2
Dim $ORDTYPE
Dim $ORD_LANG
Dim $CODE3
Dim $RACTION
Dim $VOL
Dim $COPIES
Dim $ODATE
Dim $TLOC
Dim $row1_MASTER, $acqtypecheck, $row2_MASTER, $row3_MASTER, $row4_MASTER, $row5_MASTER, $row6_MASTER, $row7_MASTER, $row8_MASTER
Dim $approval, $notification, $firm, $ICODE1_1
Dim $ITEM_STAT
Dim $BARCODE
Dim $multi_a, $multi
Dim $EPRICE_15, $15
Dim $decide
Dim $RDATE_R
Dim $FUND_R
Dim $ITEM_STAT
Dim $BIB_LOC
Dim $LABELLOC
Dim $ILOC
Dim $ADD_VOL_A, $ADD_COPY_A, $COPY_N, $VOL_N
Dim $NO_ITEM
Dim $REF_A, $REF
Dim $EBOOK_A, $EBOOK
Dim $ADD_COPY
Dim $loc_name
Dim $reserve
Dim $dean
Dim $multi_loc
Dim $T85, $T86, $T87, $T88 ;messages for text box if item needs to be bumped
Dim $RECEIVE ;receive item - today's date in rdate and save ---> 1 is received, 0 is not saved
Dim $REVIEW ;skip review of bib record ----> 1 is reviewed, 0 is skip
Dim $PRINT ;print record -----> 1 to print
Dim $BARCODE ;barcode item -----> N to skip barcode
Dim $C_INI
dim $ITYPE
Dim $FUND2
Dim $SF_NAME, $SF_590, $SF_7XX
Dim $copycat
Dim $ICODE1
Dim $advance

$advance = _LoadVar("$advance")
$copycat = _LoadVar("$copycat")
$ICODE1 = _LoadVar("$ICODE1")
;MsgBox(0, "copycat", $copycat)
;################################ MAIN ROUTINE #################################
$C_INI = _Initial()

;BEGIN LOADING VARS from isbnToTitle script
$300_E = _LoadVar("$300_E")
$300_E_E = Eval("3" & "0" & "0" & "_" & "E")

;Set messages/receive/review to default values
$RECEIVE = 1 ;default - will receive
$REVIEW = 1 ;default - will review
$T85 = ""
$T86 = ""
$T87 = ""
$T88 = ""



;Set focus on order record / Emily /
;If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;	WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Sierra record", "Please open the order record.")
;	Exit
;EndIf

;click on order record data fields
;$winsize = WinGetPos("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;_WinClick($winsize)


If WinExists("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the order record.")
	Exit
EndIf


;click on order record data fields
$winsize = WinGetPos("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
Sleep(0100)
;Workaround for Sierra for debugging
;If WinExists("[TITLE:Sierra; CLASS:SunAwtFrame]") Then
;	WinActivate("[TITLE:Sierra; CLASS:SunAwtFrame]")
;	WinWaitActive("[TITLE:Sierra; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Sierra record", "Please open the order record.")
;	Exit
;EndIf
_SendEx("^{HOME}")
Sleep(0200)
_SendEx("^a")
Sleep(0100)
_SendEx("^a")
Sleep(0100)
_SendEx("^c")
Sleep(0300)


;select all and copy order record information, parse data into array
Sleep(1000)
_DataCopy()
$ORD_REC = ClipGet()
$ORD_REC_PREP = StringRegExpReplace($ORD_REC, "[\r\n]+", "fnord")
$ORD_ARRAY_MASTER = StringSplit($ORD_REC_PREP, "fnord", 1)

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

$FUND2 = StringLeft($FUND, 3); this is for checking the fund code against the special fund list to see if we need donor notes in create item - dlc
_StoreVar("$FUND2")
;msgbox(0, "FUND2", $FUND2)

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

Run("Special Fund List.exe")
While ProcessExists("Special Fund List.exe")
	Sleep(0400)
WEnd
$SF_NAME = _LoadVar("$SF_NAME")
$SF_590 = _LoadVar("$SF_590")
$SF_7XX = _LoadVar("$SF_7XX")


_StoreVar("$UPD")
_StoreVar("$LAB_LOC")
_StoreVar("$VOL")
_StoreVar("$ICODE1")
_StoreVar("$ACCOMP")
_StoreVar("$BIB_LOC_1")
_StoreVar("$LOCAL")
_StoreVar("$dean")

If $copycat <> "1" Then


;~ ##### start rloc check #####
$RLOC = StringLeft($RLOC, 1)
Switch $RLOC
	Case "a" ;approval code
		$approval = 1
		_StoreVar("$approval")
		$ICODE1_1 = 2
		_StoreVar("$ICODE1_1")
	Case "b" ;firm order code
		$firm = 1
		$ICODE1_1 = 1
		_StoreVar("$ICODE1_1")
	Case "n"
		$notification = 1 ;notification code
		$ICODE1_1 = 1
		_StoreVar("$ICODE1_1")
	Case "c"
		$RECEIVE = 1
		$REVIEW = 0
		$NO_ITEM = 1 ;do not create item record
		$T85 = "This item is for SPECIAL COLLECTIONS."
		$T86 = "Do NOT create an item record."
		$T87 = "Do NOT barcode item."
		$T88 = "Print this record, place in item and give item to Linda."
		$PRINT = 1 ;prints record
	Case "s"
		$RECEIVE = 0
		$REVIEW = 0
		$T85 = "This item is for SERIALS."
		$T86 = "Do NOT create an item record."
		$T87 = "Do NOT barcode item."
		$T88 = "Do NOT receive this item."
		$NO_ITEM = 1
	Case "h"
		$RECEIVE = 0
		$REVIEW = 0
		$T85 = "This item is for HAMILTON."
		$T86 = "Do NOT create an item record."
		$T87 = "Do NOT barcode item."
		$T88 = "Do NOT receive this item."
		$NO_ITEM = 1
	Case "m" ;middletown order
		$ICODE1_1 = 83
		_StoreVar("$ICODE1_1")
		$firm = 1
EndSwitch
;msgbox(0, "icode", $approval)
;~ ##### end rloc check #####

;~ ##### start spec double check #####
;sometimes the order team forgets to enter the correct RLOC, but are consistant with location codes...
If $LOCATION = "spec" Then
	$RECEIVE = 1
	$REVIEW = 0
	$NO_ITEM = 1
	$T85 = "This item is for SPECIAL COLLECTIONS."
	$T86 = "Do NOT create an item record."
	$T87 = "Do NOT barcode item."
	$T88 = "Print this record, place in item and give item to Linda."
	$PRINT = 1
EndIf
;~ ##### end spec double check #####

;~ ##### start approval check #####
;make sure approval record has right codes (location & ordtype)
#cs
If $approval = 1 Then
	$ORDTYPE = StringLeft($ORDTYPE, 1)
	If $ORDTYPE <> "a" Or $LOCATION <> "acq" Then
		MsgBox(64, "Approval Code Match? - Macro Stop", "It appears that the order record indicates Approval, but does not have the correct codes." & @CR & "Please correct the codes and run the macro again.")
		Exit
	EndIf
EndIf
;~ ##### end approval check #####
#ce

;~ ##### start ref check #####
#cs
	$REF_A = _ArraySearch($ORD_ARRAY_MASTER, "Ref", 0, 0, 0, 1)
	old way of checking, caused false positives
	as of 2010, this check will be made using the location code
#ce
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
;~ ##### end ref check #####

;~ ##### start notification check #####
;make sure notification record has right codes (ordtype). Also bumps for certain ordtype codes.
If $notification = 1 Then
	$ORDTYPE = StringLeft($ORDTYPE, 1)
	Switch $ORDTYPE
		Case "u", "s", "p" ;added volumes!
			$RECEIVE = 1
			$REVIEW = 0
			$ITEM_STAT = "i"
			$LABELLOC = ""
			$T85 = "Added volume."
			$T88 = "BUMP to Susan Keyes"
		Case "t", "n" ;appropriate matching codes for notification orders
			Sleep(0010)
		Case Else
			MsgBox(64, "notification Code Match? - Macro Stop", "It appears that the order record indicates Notification, but does not have the correct codes." & @CR & "Please correct the codes and run the macro again.")
			Exit
	EndSwitch
EndIf
;~ ##### end notification check #####

;~ ##### start multi fund check #####
;some items are bought with multiple funds. This gets these funds into a multi variable for later use
If $FUND = "multi" Then
	$multi = _arrayItemString($ORD_ARRAY_MASTER, "FUNDS	")
	$multi = StringTrimLeft($multi, 6)
EndIf
;~ ##### end multi fund check #####

;~ ##### start Middletown DVD check #####
$FORM = StringLeft($FORM, 1)
If $FORM = "j" Then
	Switch $LOCATION
		Case "mdli", "mdr", "mdim", "mdju"
			$BARCODE = "N"
			$T85 = "Middletown DVD"
			$T88 = "Send to Middletown DVD cataloger."
		Case Else
			$BARCODE = "N" ;don't barcode DVDs in general
	EndSwitch
EndIf
;~ ##### end Middletown material check #####

;~ ##### start e price checking #####
;eprices are checked for notification and firm orders
;15 is added to the base price. If the price on the item slip is beyond base + 15, then bump to Acq Librarian
If $notification = 1 Or $firm = 1 Then
	$EPRICE = StringTrimLeft($EPRICE, 1)
	$15 = 15
	$EPRICE_15 = $EPRICE + $15
	$decide = MsgBox(4, "E Price Check", "Is the net price betweem " & $EPRICE & " and " & $EPRICE_15 & "?")
	Switch $decide
		Case 6 ;price falls within range
		Case 7 ;price does not fall within range
			MsgBox(64, "Price does not match - Macro Stop", "Do not receive this item." & @CR & "Ask Karen Clift for approval of price.")
			Exit
	EndSwitch
EndIf
;~ ##### end e price checking #####

;~ ##### start volume/copy check #####
If $VOL > 1 Or $COPIES > 1 Then
	$decide = MsgBox(4, "Copies Check", "According to this order record, " & @CR & "we are expecting " & $VOL & " volumes and " & $COPIES & "copies of this title." & @CR & "Have we received all of them with this shipment?")
	Switch $decide
		Case 6 ;all copies/vols received
			;WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
			WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
			Sleep(0100)
			_SendEx("^{HOME}")
			Sleep(0100)
			_SendEx("{ENTER}z")
			Sleep(0100)
			If $VOL > 1 Then
				_SendEx("Received volumes{SPACE}")
				$VOL_N = InputBox("Volume #", "Enter the number of volumes in the internal note and close this box to proceed." & @CR & "Please enter the volume designation for all the volumes received in this shipment (for example, '1-2' OR '5, 7-8' OR '3 and appendix')")
				Sleep(0100)
				_SendEx($VOL_N & ". " & $C_INI)
			Else
				_SendEx("Received ")
				$COPY_N = InputBox("Copy #", "Enter the number of copies.", "")
				_SendEx("{SPACE}" & $COPY_N & "{SPACE}copies. " & $C_INI)
			EndIf
		Case 7 ;not all copies/vols received
			;WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
			WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
			Sleep(0100)
			_SendEx("^{HOME}")
			Sleep(0100)
			_SendEx("{ENTER}z")
			Sleep(0100)
			If $VOL > 1 Then
				_SendEx("Received volumes{SPACE}")
				$VOL_N = InputBox("Volume #", "Enter the number of volumes in the internal note and close this box to proceed." & @CR & "Please enter the volume designation for all the volumes received in this shipment (for example, '1-2' OR '5, 7-8' OR '3 and appendix')")
				Sleep(0100)
				_SendEx($VOL_N & ". " & $C_INI)
			Else
				_SendEx("Received ")
				$COPY_N = InputBox("Copy #", "Enter the number of copies.", "")
				_SendEx("{SPACE}" & $COPY_N & "{SPACE}copies. " & $C_INI)
			EndIf
			_SendEx("^s");save record at this point
			Sleep(0100)
			WinWait("[TITLE:WARNING; CLASS:SunAwtDialog]", "")
			_SendEx("{ENTER}")
	EndSwitch
EndIf
;~ ##### end volume check #####

#cs
;~ ##### start RDATE check #####
$RDATE_R = StringInStr($RDATE, "-  -")
If $RDATE_R = 0 Then
	$decide = MsgBox(4, "Already received", "This item was already received on " & $RDATE & ". Do you want to override the RDATE?")
	Switch $decide
		Case 6
		Case 7 ;don't override the rdate - will exit the script
			Exit
	EndSwitch
EndIf
;~ ##### End RDATE check #####
#ce

;~ ##### start local use check #####
;used to determine BCODE3 and ITYPE codes
$LOCAL = _ArraySearch($ORD_ARRAY_MASTER, "LOCAL USE ONLY", 0, 0, 0, 1)
;~ ##### end local use check #####

;~ ##### start textbook fund check #####
$FUND_R = StringInStr($FUND, "4a Textbooks")
If $FUND_R > 0 Then
	$T85 = "This item is for TEXTBOOK RESERVES."
	$T86 = "Create an item record."
	$RECEIVE = 1
	$REVIEW = 0
EndIf
;~ ##### end textbook fund check #####

;~ ##### start status check #####
$STATUS = StringLeft($STATUS, 1)
Switch $STATUS
	Case "z"
		$RECEIVE = 0
		$REVIEW = 1
		$T85 = "The approval has been rejected."
		$T86 = "Do NOT create an item record."
		$T87 = "Do NOT barcode item."
		$T88 = "Please see internal notes."
		$BARCODE = "N"
		$NO_ITEM = 1
	Case "a" ;move along, nothing to see here
		$RECEIVE = 1
	Case "q"
		$RECEIVE = 0
		$REVIEW = 0
		$T85 = "This item has been partially paid."
	Case "g", "c", "d", "e", "f" ;serials order
		$RECEIVE = 0
		$REVIEW = 0
		$T85 = "This is a serial order."
		$T86 = "Give item to Serials team."
		$T87 = "Do NOT create an item record."
		$T88 = "Do NOT barcode item."
		$BARCODE = "N"
		$NO_ITEM = 1
EndSwitch
;~ ##### end status check #####

;~ ##### start Score form check #####
;Music scores go to conservation and then to Music Librarian for Cataloging
If $FORM = "w" Then
	$REVIEW = 0
	$BARCODE = "N"
	$LOCATION = "muli"
	$ITYPE = 12
	$ITEM_STAT = "i"
	$T85 = "Do NOT attach barcode."
	$T86 = "SEND to Conservation."
	;$score = 1
EndIf
;~ ##### end form check #####
;MsgBox(0, "variables", $ITYPE & $LOCATION)

;~ ##### start MLC check #####
;MLC = Music Listening Center
If $LOCATION = "murcd" Then
	$RECEIVE = 1
	$REVIEW = 0
	$PRINT = 1
	$T85 = "Do NOT attach item record."
	$T86 = "Do NOT attach barcode."
	$T87 = "Print off record and tape to CD."
	$T88 = "SEND TO MUSIC LIBRARY LISTENING CENTER ASSOCIATE."
	$BARCODE = "N"
	$NO_ITEM = 1
EndIf
;~ ##### END MLC check #####

;~ ##### start video games check #####
If $LOCATION = "game1" Then
	$REVIEW = 0
	$T85 = "Do NOT attach barcode."
	$T88 = "SEND to Cataloging Librarian."
	$BARCODE = "N"
EndIf
;~ ##### END video games check #####

;~ ##### start dean's material check #####
If $LOCATION = "dean" Then
	$dean = 1
EndIf
;~ ##### end dean's material check #####

;~ ##### start imc check #####
If $LOCATION = "imdvd" Or $LOCATION = "imvc" Or $LOCATION = "imm" Then
	$REVIEW = 0
	$T85 = "Do NOT attach barcode."
	$T88 = "SEND to Marilee."
	$BARCODE = "N"
EndIf
;~ ##### END imc check #####

;~ ##### start added volume/copy check #####
$ADD_VOL_A = _ArraySearch($ORD_ARRAY_MASTER, "ADDED VOL", 0, 0, 0, 1)
$ADD_COPY_A = _ArraySearch($ORD_ARRAY_MASTER, "ADDED COP", 0, 0, 0, 1)

If $ADD_VOL_A > -1 Or $ADD_COPY_A > -1 Then
	$ADD_COPY = 1
Else
	$ADD_COPY = 0
EndIf

If $ADD_COPY = 1 And $LOCATION <> "mdli" And $LOCATION <> "mdr" And $LOCATION <> "mdim" And $LOCATION <> "mdju" Then
	$RECEIVE = 1
	$REVIEW = 0
	$ITEM_STAT = "i"
	$LOCATION = "none"
	$T85 = "Added volume/copy"
	$T88 = "BUMP TO Susan Keyes. Videos go to Marilee."
EndIf
;~ ##### end added volume/copy check #####

;~ ##### Start ebook check #####
;sets variable that is used once item record is created
$EBOOK_A = _ArraySearch($ORD_ARRAY_MASTER, "Keep paperwork with book", 0, 0, 0, 1)
If $EBOOK_A > -1 Then
	$EBOOK = 1
EndIf
;~ ##### end ebook check #####

Dim $multi_loc_e

;~ ##### Start multi location check #####
If $LOCATION = "multi" Then
	$multi_loc = _arrayItemString($ORD_ARRAY_MASTER, "LOCATIONS	")
	$multi_loc = StringTrimLeft($multi_loc, 10)
EndIf
;~ ##### end multi location check #####

;~ ##### Start reserve check #####
#cs
	Items for reserve should not have the reserve location code in the item record for
	items that are still in TS. Circulation is responsible for changing the item location
	code from the general collection code to the reserves code.
	The script enters an internal note stating to circ to insert general collection iloc
	when the item is taken off reserve.
#ce
If $LOCATION = "aars" Or $LOCATION = "docrs" Or $LOCATION = "imrs" Or $LOCATION = "kngrs" Or $LOCATION = "murs" Or $LOCATION = "scrs" Then
	;WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	Sleep(0100)
	_SendEx("^{HOME}")
	Sleep(0100)
	_SendEx("{ENTER}z")
	Sleep(0100)
	_SendEx("Rush for{SPACE}")
	Sleep(0100)
	Switch $LOCATION
		Case "aars"
			$LOCATION = "aali"
			$loc_name = "ArtArch"
		Case "docrs"
			$LOCATION = "docli"
			$loc_name = "Gov Docs"
		Case "imrs"
			$LOCATION = "imcbk"
			$loc_name = "IMC"
		Case "kngrs"
			$LOCATION = "kngli"
			$loc_name = "King"
		Case "murs"
			$LOCATION = "muli"
			$loc_name = "Music"
		Case "scrs"
			$LOCATION = "scli"
			$loc_name = "Science"
	EndSwitch
	_SendEx($loc_name & "{SPACE}reserves. Place in iloc " & $LOCATION & " when taken off of reserve.")
	Sleep(0500)
	_SendEx("^s")
	Sleep(0100)
	WinWait("[TITLE:WARNING; CLASS:SunAwtDialog]", "")
	_SendEx("{ENTER}")
	$reserve = 1
EndIf
;~ ##### End reserve check #####

;~ ##### BEGIN DISPLAY BOX #####
;If the item needs to be bumped, the display box will state why it's being bumped and to who
If $T85 <> "" Or $T86 <> "" Or $T87 <> "" Or $T88 <> "" Then
	MsgBox(0, "Receipt Cataloging", $T85 & @CR & $T86 & @CR & $T87 & @CR & $T88)
EndIf
;~ ##### END DISPLAY BOX #####

;~ ##### BEGIN RECEIVING ITEM #####
If $RECEIVE = 1 Then
	;WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	;WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	_SendEx("^{HOME}")
	Sleep(0100)
	_SendEx("{TAB 4}")
	Sleep(0100)
	;$RDATE_R = StringInStr($RDATE, "-  -")
	;if $rdate_r <> 0 Then
	_SendEx("t");enter today's date for rdate
	;EndIf
	If $300_E = 0 Or $300_E_E = "" Then ;accompanying material note check - triggered by variable set in isbnToTitle
		Sleep(0100)
	Else
		$decide = MsgBox(4, "Accompanying Material", "This item should have the following accompanying material: " & @CR & $300_E & @CR & "Does the item have the listed material(s)?")
		Switch $decide
			Case 6
				_SendEx("^{END}")
				Sleep(0100)
				_SendEx("{ENTER}z")
				Sleep(0100)
				_SendEx("Rec'd " & $300_E & " with item. " & $C_INI)
			Case 7
				MsgBox(64, "Missing - Macro Stop", "Item is missing accompanying material. Macro will stop.")
				Exit
		EndSwitch
	EndIf
	Sleep(0100)
	_SendEx("^s");save
	WinWait("[TITLE:WARNING; CLASS:SunAwtDialog]", "")
	_SendEx("{ENTER}")
EndIf

;~ ##### BEGIN PRINT RECORD #####
If $PRINT = 1 Then
	;WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	;WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	_SendEx("^p")
	MsgBox(0, "Print record", "The record has been printed." & @CR & "Attach print-out to item.")
EndIf

;START STORING VARIABLES

_StoreVar("$approval")
_StoreVar("$BARCODE")
_StoreVar("$BIB_LOC")
_StoreVar("$COPIES")
_StoreVar("$firm")
_StoreVar("$FUND")
_StoreVar("$ICODE1_1")
_StoreVar("$ITEM_STAT")
_StoreVar("$LABELLOC")
_StoreVar("$LOCATION")
_StoreVar("$multi")
_StoreVar("$notification")
_StoreVar("$RLOC")
_StoreVar("$STATUS")
_StoreVar("$T85")
_StoreVar("$T86")
_StoreVar("$T87")
_StoreVar("$T88")
_StoreVar("$VOL")

_StoreVar("$LOCAL")
_StoreVar("$EBOOK")
_StoreVar("$reserve")
_StoreVar("$dean")
_StoreVar("$multi_loc")
_StoreVar("$ITYPE")
_StoreVar("$FORM")

;MsgBox(0, "variables", $RLOC & $EPRICE & $LOCATION)

;~ ##### BEGIN SKIP REVIEW / CREATE ITEM RECORD #####
If $REVIEW = 0 Or $LOCATION = "mdli" Or $LOCATION = "mdr" Or $LOCATION = "mdim" Or $LOCATION = "mdju" Or $LOCATION = "mdrs" Then
	Run("ILOC.exe") ;grab item code information related to location code
	While ProcessExists("ILOC.exe")
		Sleep(1000)
	WEnd
	;WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	;WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinActivate("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
EndIf


If $NO_ITEM = 1 Then
	_ClearBuffer()
	Exit
ElseIf $REVIEW = 0 Or $LOCATION = "mdli" Or $LOCATION = "mdr" Or $LOCATION = "mdim" Or $LOCATION = "mdju" Or $LOCATION = "mdrs" Then
	Sleep(0400)
	$nondlc = 1
	_StoreVar("$nondlc")
	Run("ItemCreate");creates item record
Else
	$ccat = 1
	_StoreVar("$ccat")
	MsgBox(0, "Order notes", "Please make a mental note of any relevant order record notes!")
	Run("BibRecord.exe");moves onto bib record function
EndIf


EndIf
Exit
