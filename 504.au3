#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\..\Programming\Macros\withdrawals\Images\autoiticon.ico
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

 Name of script: 504

 Script Function:
	This script guides the entry of bib ref/index notes into a 504 field in both
	Sierra and OCLC depending on what the user wants.

 Programs used: III Sierra 1.2.1_7, Windows 8.1, OCCL Connexion

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
TraySetIcon(@WorkingDir & "\images\blue5.ico")

;######### DECLARE VARIABLES #########
Dim $PG_NUM, $INDEX, $BIB_REF
Dim $choose

;################################ MAIN ROUTINE #################################
;focus on Connexion screen
$choose = InputBox("Sierra or OCLC?", "Enter s for Sierra or o for OCLC")
if $choose = "o" Then
	If WinExists("OCLC Connexion") Then
		WinActivate("OCLC Connexion")
	Else
		MsgBox(64, "nope", "Please open and log into Connexion. Click ok when you are logged in")
	EndIf



;Start prompts for information about page numbers and index
$BIB_REF = MsgBox(4, "Bib ref?", "Does this item have bibliographic references?")
If $BIB_REF = 6 Then

$PG_NUM = InputBox("Page Numbers?", "Enter page number(s) of bibliography, if appropriate.", "")
$BIB_REF = " bibliographic references"
Else
	$BIB_REF = ""
EndIf
$INDEX = MsgBox(4, "Index?", "Does the item have an index?")
If $INDEX = 7 AND $BIB_REF = "" Then
	Exit
Else

Sleep(0100)
EndIf

;Send line to Connexion client record

_SendEx("{ENTER}504{SPACE 2}Includes"& $BIB_REF)
If $PG_NUM <> "" Then
	_SendEx("{SPACE}(pages "& $PG_NUM &")")
EndIf
If $INDEX = 6 Then
	If $BIB_REF = "" Then
		_SendEx("{SPACE}index")
	Else

	_SendEx("{SPACE}and index")
	EndIf
EndIf
_SendEx(".")

	ElseIf $choose = "s" Then
WinActivate("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")

	$BIB_REF = MsgBox(4, "Bib ref?", "Does this item have bibliographic references?")
If $BIB_REF = 6 Then

$PG_NUM = InputBox("Page Numbers?", "Enter page number(s) of bibliography, if appropriate.", "")
$BIB_REF = " bibliographic references"
Else
	$BIB_REF = ""
EndIf
$INDEX = MsgBox(4, "Index?", "Does the item have an index?")
If $INDEX = 7 AND $BIB_REF = "" Then
	Exit
Else

Sleep(0100)
EndIf


_SendEx("{ENTER}n504{SPACE 2}Includes"& $BIB_REF)
If $PG_NUM <> "" Then
	_SendEx("{SPACE}(pages "& $PG_NUM &")")
EndIf
If $INDEX = 6 Then
	If $BIB_REF = "" Then
		_SendEx("{SPACE}index")
	Else

	_SendEx("{SPACE}and index")
	EndIf

EndIf
	EndIf

MsgBox(0, "MARC 008", "Don't forget to adjust the fields in the 008!")
