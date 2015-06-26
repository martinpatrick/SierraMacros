#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=images\autoiticon.ico
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

 Name of script: 949

 Script Function:
	This script enters information into OCLC Connexion to enable exporting and
	overlaying of bib records. It pulls the bib record # from Sierra.

 Programs used: III Sierra 1.2.1_7, Windows 8.1, OCLC Connexion

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
TraySetIcon(@WorkingDir & "\Images\red9.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_NUM
;Dim $BIB_LOC_1
Dim $OLOCATION

;$BIB_LOC_1 = _LoadVar("$BIB_LOC_1")
$OLOCATION = _LoadVar("$OLOCATION")

;################################ MAIN ROUTINE #################################
;focus on Millennium bib record screen
;If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Millennium record", "Please open the bib record.")
;	Exit
;EndIf

;/Emily/

#cs
If WinExists("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the bib record.")
	Exit
EndIf
#ce


If WinExists("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the bib record.")
	Exit
EndIf


;Get bib number from window title
;$BIB_NUM = WinGetTitle("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;/Emily/
;$BIB_NUM = WinGetTitle("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]")

$BIB_NUM = WinGetTitle("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")

_SendEx("!{F4}")

;focus on Connexion
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")

;send 949 line
_SendEx("^{END}")
Sleep(0100)
_SendEx("{ENTER}949{SPACE 2}*recs-b;b3--;ov-." & $BIB_NUM & ';bn=' & $OLOCATION)
