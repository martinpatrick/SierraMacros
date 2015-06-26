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

 Name of script: ILOC

 Script Function:
	This script decides what the full codes related to location and use
	should be for the final item record.

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
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $T25
Dim $LOCATION
Dim $BIB_LOC
Dim $LABELLOC
Dim $049
Dim $ITYPE
Dim $ITEM_STAT
Dim $BIB_LOC_L1
Dim $BIB_LOC_L2
Dim $REF
dim $score
Dim $FORM
Dim $BIB_LOC_1

;################################ MAIN ROUTINE #################################
$LOCATION = _LoadVar("$LOCATION") ;determined in copy cat/receipt scripts
$score = _LoadVar("$score") ; needed to make the item record have the correct itype and location
$FORM = _LoadVar("$FORM")

msgbox(0, "ILOC", "ILOC is running")
$T25 = 0

Switch $LOCATION

	Case "none"
		Sleep(0100)

; ART/ARCH
	Case "aali"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch"
		$049 = "MIAH"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "aarm"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch CD-ROM"
		$049 = "MIAH"
		$ITYPE = 21
		$ITEM_STAT = "a"
	Case "aac"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch Circ"
		$049 = "n/a"
		$ITYPE = 18
		$ITEM_STAT = "a"
	Case "aacls"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch"
		$049 = "MIAH"
		$ITYPE = 8
		$ITEM_STAT = "a"
	Case "aafo"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch Folio"
		$049 = "MIAH"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "aar"
		$BIB_LOC = "multi"
		$BIB_LOC_L1 = "aal"
		$BIB_LOC_L2 = "ref"
		$LABELLOC = "ArtArch Ref"
		$049 = "MIAH"
		$ITYPE = 21
		$ITEM_STAT = "o"
		$REF = 1
	Case "aap"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch Per"
		$049 = "MIJ&"
		$ITYPE = 31
		$ITEM_STAT = "g"
	Case "aast"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch"
		$049 = "MIAH"
		$ITYPE = 0
		$ITEM_STAT = "f"
	Case "aars"
		$LOCATION = "acq"
		$T25 = "RESERVE ITEM - GIVE TO PROCESSING"
	Case "aavc"
		$BIB_LOC = "aal"
		$LABELLOC = "ArtArch VideoC"
		$049 = "MIAH"
		$ITYPE = 8
		$ITEM_STAT = "a"

;GOV DOCS
	Case "docli"
		$BIB_LOC = "doc"
		$LABELLOC = ""
		$049 = "MIAM"
		$ITYPE = 19
		$ITEM_STAT = "-"
	Case "docrm"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov CD-ROM"
		$049 = "MIB+"
		$ITYPE = 21
		$ITEM_STAT = "a"
	Case "docc"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov Census"
		$049 = "MIBD"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "doceu"
		$BIB_LOC = "doc"
		$LABELLOC = "GovEuro"
		$049 = "MIKA"
		$ITYPE = 19
		$ITEM_STAT = "-"
	Case "docl"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov Law"
		$049 = "MIBI"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "doco"
		$BIB_LOC = "doc"
		$LABELLOC = ""
		$049 = "MIAM"
		$ITYPE = 19
		$ITEM_STAT = "-"
	Case "docr"
		$BIB_LOC = "multi"
		$BIB_LOC_L1 = "doc"
		$BIB_LOC_L2 = "ref"
		$LABELLOC = "Gov   Ref"
		$049 = "MIAD"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "docun"
		$BIB_LOC = "doc"
		$LABELLOC = "GovUN"
		$049 = "MIKA"
		$ITYPE = 19
		$ITEM_STAT = "-"
	Case "docfc"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov M'fiche"
		$049 = "MIAJ"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "docfm"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov M'film"
		$049 = "MIAL"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "docsf"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov USMicro"
		$049 = "MIB"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "docsr"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov USRef"
		$049 = "MIB+"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "docrs"
		$BIB_LOC = "doc"
		$LABELLOC = "Gov   Res"
		$049 = "n/a"
		$ITYPE = 8
		$ITEM_STAT = "-"
;No external collections...

;IMC
	Case "imbk"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC   Book"
		$049 = "MIB+"
		$ITYPE = 22
		$ITEM_STAT = "-"
	Case "imci"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC  Circ"
		$049 = "MIBN"
		$ITYPE = 18
		$ITEM_STAT = "a"
	Case "imdvd"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC   DVD"
		$049 = "MIBn"
		$ITYPE = 2
		$ITEM_STAT = "-"

	Case "imcur"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC CurrMat"
		$049 = "MIBN"
		$ITYPE = 2
		$ITEM_STAT = "-"
	Case "imfs"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC  Films"
		$049 = "MIBN"
		$ITYPE = 2
		$ITEM_STAT = "-"
	Case "imm"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC CurrMat"
		$049 = "MIBN"
		$ITYPE = 2
		$ITEM_STAT = "y"
	Case "imrs"
		$LOCATION = "acq"
	Case "imbkt"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC Textbook"
		$049 = "MIBN"
		$ITYPE = 22
		$ITEM_STAT = "-"
	Case "imvc"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC VideoC"
		$049 = "MIBn"
		$ITYPE = 2
		$ITEM_STAT = "-"

	Case "imvr"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC Video Clsd"
		$049 = "MIBN"
		$ITYPE = 8
		$ITEM_STAT = "a"
	Case "imtst"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC Test"
		$049 = "n/a"
		$ITYPE = 23
		$ITEM_STAT = "-"
	Case "imje"
		$BIB_LOC = "imc"
		$LABELLOC = "Juv  Easy"
		$049 = "MIBN"
		$ITYPE = 22
		$ITEM_STAT = "-"
	Case "imjb"
		$BIB_LOC = "imc"
		$LABELLOC = "Juv BigBook"
		$049 = "MIBN"
		$ITYPE = 22
		$ITEM_STAT = "-"
	Case "imju"
		$BIB_LOC = "imc"
		$LABELLOC = "Juv"
		$049 = "MIBN"
		$ITYPE = 22
		$ITEM_STAT = "-"
	Case "imrdr"
		$BIB_LOC = "imc"
		$LABELLOC = "Juv Reader"
		$049 = "MIBN"
		$ITYPE = 22
		$ITEM_STAT = "-"
	Case "imlng"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC LangMat"
		$049 = "MIBN"
		$ITYPE = 22
		$ITEM_STAT = "-"
	Case "kngln"
		$BIB_LOC = "imc"
		$LABELLOC = "Now in King"
		$049 = "MIJP"
		$ITYPE = 31
		$ITEM_STAT = "o"
	Case "imp"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC   Per"
		$049 = "MIJQ"
		$ITYPE = 31
		$ITEM_STAT = "g"
	Case "imcra"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'Card"
		$049 = "MIAW"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imcrc"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'Card"
		$049 = "MIJG"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imcrp"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'Card Per"
		$049 = "MIJH"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imfca"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'fiche"
		$049 = "MIAI"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imfcc"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'fiche"
		$049 = "MIJI"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imfcn"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'fiche NewP"
		$049 = "MIJJ"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imfcp"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'fiche Per"
		$049 = "MIJK"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imfma"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'film"
		$049 = "MIAD"
		$ITYPE = 21
		$ITEM_STAT = "q"
	Case "imfmc"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'film"
		$049 = "MIJM"
		$ITYPE = 21
		$ITEM_STAT = "q"
	Case "imfmp"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'film Per"
		$049 = "MIJN"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "impra"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'print"
		$049 = "MIJO"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "imprc"
		$BIB_LOC = "imc"
		$LABELLOC = "IMC M'print"
		$049 = "MIJO"
		$ITYPE = 21
		$ITEM_STAT = "o"

;KING
	Case "cat"
		$BIB_LOC = "kngl"
		$LABELLOC = "Cat"
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "kngli"
		$BIB_LOC = "kngl"
		$LABELLOC = ""
		$049 = "MIAA"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "kngrm"
		$BIB_LOC = "kngl"
		$LABELLOC = "King CD-ROM"
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "a"
	Case "kngcm"
		$BIB_LOC = "kngl"
		$LABELLOC = "King CIM"
		$049 = "MIAA"
		$ITYPE = 18
		$ITEM_STAT = "a"
	Case "kngc"
		$BIB_LOC = "kngl"
		$LABELLOC = "King Circ"
		$049 = "MIAA"
		$ITYPE = 18
		$ITEM_STAT = "a"
	Case "kngfo"
		$BIB_LOC = "kngl"
		$LABELLOC = "King Folio"
		$049 = "MIAA"
		$ITYPE = 0
		$ITEM_STAT = "a"
	Case "kngin"
		$BIB_LOC = "kngl"
		$LABELLOC = "King Index"
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "kngr"
		$BIB_LOC = "multi"
		$BIB_LOC_L1 = "kngl"
		$BIB_LOC_L2 = "ref"
		$LABELLOC = "King  Ref"
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "o"
		$REF = 1
	Case "cdlan"
		$BIB_LOC = "kngl"
		$LABELLOC = "CD-ROM LAN"
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "a"
	Case "kngln"
		$BIB_LOC = "kngl"
		$LABELLOC = "King Newsp"
		$049 = "MIJP"
		$ITYPE = 8
		$ITEM_STAT = "-"
	Case "kngp"
		$BIB_LOC = "kngl"
		$LABELLOC = "King Per"
		$049 = "MIA7"
		$ITYPE = 31
		$ITEM_STAT = "g"
	Case "kngle"
		$BIB_LOC = "kngl"
		$LABELLOC = "Leisure"
		$049 = "MIAA"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "kngfc"
		$BIB_LOC = "kngl"
		$LABELLOC = "King M'fiche"
		$049 = "MIBJ"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "kngfm"
		$BIB_LOC = "kngl"
		$LABELLOC = "King M'film"
		$049 = "MIBK"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "kngrs"
		$LOCATION = "acq"
		$T25 = "RESERVE ITEM - GIVE TO PROCESSING"

;MUSIC
	Case "muli"
		If $score = 1 Then
		 $BIB_LOC = "mul"
		$LABELLOC = "Music"
		$049 = "MIAE"
		$ITYPE = 12
		$ITEM_STAT = "-"
	 Else
		$BIB_LOC = "mul"
		$LABELLOC = "Music"
		$049 = "MIAE"
		$ITYPE = 0
		$ITEM_STAT = "-"
	 EndIf
	Case "mufo"
	  if $score = 1 Then
	    $BIB_LOC = "mul"
		$LABELLOC = "Music Folio"
		$049 = "MIAE"
		$ITYPE = 12
		$ITEM_STAT = "-"
	 Else
		$BIB_LOC = "mul"
		$LABELLOC = "Music Folio"
		$049 = "MIAE"
		$ITYPE = 0
		$ITEM_STAT = "-"
	 EndIf

	Case "mudc"
		$BIB_LOC = "mul"
		$LABELLOC = "Music"
		$049 = "MIAE"
		$ITYPE = 8
		$ITEM_STAT = "-"
	Case "muc"
		$BIB_LOC = "mul"
		$LABELLOC = "Music Circ"
		$049 = "MIAE"
		$ITYPE = 18
		$ITEM_STAT = "a"
	Case "mur"
		$BIB_LOC = "multi"
		$BIB_LOC_L1 = "mul"
		$BIB_LOC_L2 = "ref"
		$LABELLOC = "Music Ref"
		$049 = "MIAE"
		$ITYPE = 21
		$ITEM_STAT = "o"
		$REF = 1
	Case "murs"
		$LOCATION = "acq"
		$T25 = "RESERVE ITEM - GIVE TO PROCESSING"
	Case "mup"
		$BIB_LOC = "mul"
		$LABELLOC = "Music Per"
		$049 = "MIJV"
		$ITYPE = 31
		$ITEM_STAT = "g"
	Case "mufm"
		$BIB_LOC = "mul"
		$LABELLOC = "Music M'film"
		$049 = "MIBC"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "murc"
		$BIB_LOC = "mul"
		$LABELLOC = "Music"
		$049 = "MIAE"
		$ITYPE = 1
		$ITEM_STAT = "-"
	Case "murcv"
		$BIB_LOC = "mul"
		$LABELLOC = ""
		$049 = "MIJU"
		$ITYPE = 1
		$ITEM_STAT = "o"
	Case "murcd"
		$BIB_LOC = "mul"
		$LABELLOC = "Music Ref"
		$049 = "MIJU"
		$ITYPE = 1
		$ITEM_STAT = "o"


;SCIENCE
	Case "scli"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci"
		$049 = "MIAS"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "scat"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci Atlas"
		$049 = "MIAS"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scc"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci Circ"
		$049 = "MIAS"
		$ITYPE = 18
		$ITEM_STAT = "a"
	Case "scin"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci Index"
		$049 = "MIAS"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scrr"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci ReadyRef"
		$049 = "MIAS"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scr"
		$BIB_LOC = "multi"
		$BIB_LOC_L1 = "scl"
		$BIB_LOC_L2 = "ref"
		$LABELLOC = "Sci   Ref"
		$049 = "MIAS"
		$ITYPE = 21
		$ITEM_STAT = "o"
		$REF = 1
	Case "scmp"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci   Map"
		$049 = "MIBE"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scmpc"
		$BIB_LOC = "scl"
		$LABELLOC = "SciMapCb"
		$049 = "MIAS"
		$ITYPE = 21
		$ITEM_STAT = "0"
	Case "scmph"
		$BIB_LOC = "scl"
		$LABELLOC = "SciMapHz"
		$049 = "MIBE"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scmpv"
		$BIB_LOC = "scl"
		$LABELLOC = "SciMapVt"
		$049 = "MIBE"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scp"
		$BIB_LOC = "scl"
		$LABELLOC = "scp"
		$049 = "MIA8"
		$ITYPE = 31
		$ITEM_STAT = "g"
	Case "sccrp"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci M'card Per"
		$049 = "MIJI"
		$ITYPE = 31
		$ITEM_STAT = "g"
	Case "scfcp"
		$BIB_LOC = "scl"
		$LABELLOC = "scp M'fiche Per"
		$049 = "MIJ3"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scfmp"
		$BIB_LOC = "scl"
		$LABELLOC = "scp M'film Per"
		$049 = "MIJ5"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "sccra"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci M'Card"
		$049 = "MIBH"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "sccrc"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci M'card"
		$049 = "MIJZ"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scfca"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci M'fiche"
		$049 = "MIBF"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scfcc"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci M'fiche"
		$049 = "MIJ2"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scfma"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci M'film"
		$049 = "MIBG"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scfmc"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci M'film"
		$049 = "MIJ4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scfo"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci Folio"
		$049 = "MIAS"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "sckuc"
		$BIB_LOC = "scl"
		$LABELLOC = "Sci Kuchler"
		$049 = "MIJY"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "scrs"
		$LOCATION = "acq"
		$T25 = "RESERVE ITEM - GIVE TO PROCESSING"

;ARCHIVES/SPEC
	Case "arcli"
		$BIB_LOC = "spe"
		$LABELLOC = "Archives"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "speco"
		$BIB_LOC = "spe"
		$LABELLOC = "Cov"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spemy"
		$BIB_LOC = "spe"
		$LABELLOC = "Myaamia"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spec"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec (LC)"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "specd"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec (Dewey)"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spcfo"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec Folio"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "speck"
		$BIB_LOC = "spe"
		$LABELLOC = "SpecKing"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spkfo"
		$BIB_LOC = "spe"
		$LABELLOC = "SpecKing Folio"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "specv"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec Coll Closed Stks"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spevf"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec Coll Closed Stks folio"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spmss"
		$BIB_LOC = "spe"
		$LABELLOC = "SpecMSS"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "specs"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec SchoolBk"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spe18"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec1841"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "specp"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec Per"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spefc"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec M'fiche"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"
	Case "spefm"
		$BIB_LOC = "spe"
		$LABELLOC = "Spec M'film"
		$049 = "MIA4"
		$ITYPE = 21
		$ITEM_STAT = "o"


;SWORD
	Case "rgspa"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIJ&"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rgspc"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA4"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rgspd"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAM"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rgsph"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "OHMM"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rgspk"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA7"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rgspm"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIJV"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rgsps"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA8"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rgsta"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAH"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rgstc"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA4"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rgstd"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAM"
		$ITYPE = 19
		$ITEM_STAT = "k"
	Case "rgstg"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "OMMM"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rgsth"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "OHMM"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rgsti"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIbn"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rgstk"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAA"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rgstm"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAE"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rgstr"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "n/a"
		$ITYPE = 21
		$ITEM_STAT = "k"
	Case "rgsts"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAS"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2pa"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIJ&"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rg2pc"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA4"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rg2pd"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAM"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rg2ph"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "OHMM"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rg2pk"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA7"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rg2pm"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIJV"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rg2ps"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA8"
		$ITYPE = 31
		$ITEM_STAT = "k"
	Case "rg2ta"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAH"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2tc"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIA4"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2td"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAM"
		$ITYPE = 19
		$ITEM_STAT = "k"
	Case "rg2tl"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAH"
		$ITYPE = 21
		$ITEM_STAT = "k"
	Case "rg2tg"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAM"
		$ITYPE = 21
		$ITEM_STAT = "k"
	Case "rg2tf"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIJL"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2th"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "OHMM"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2ti"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIBN"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2tk"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAA"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2tm"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAE"
		$ITYPE = 0
		$ITEM_STAT = "k"
	Case "rg2tr"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "n/a"
		$ITYPE = 21
		$ITEM_STAT = "k"
	Case "rg2ts"
		$BIB_LOC = "swdep"
		$LABELLOC = "RegiStor"
		$049 = "MIAM"
		$ITYPE = 0
		$ITEM_STAT = "k"

;ONLINE
	Case "doclo"
		$BIB_LOC = "onl"
		$LABELLOC = ""
		$049 = "MIAM"
		$ITYPE = 21
		$ITEM_STAT = "h"
	Case "onlm"
		$BIB_LOC = "onl"
		$LABELLOC = ""
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "h"
	Case "onlo"
		$BIB_LOC = "onl"
		$LABELLOC = ""
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "h"
	Case "onli"
		$BIB_LOC = "onl"
		$LABELLOC = ""
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "h"
	Case "onlw"
		$BIB_LOC = "onl"
		$LABELLOC = ""
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "h"
	Case "onl$"
		$BIB_LOC = "onl"
		$LABELLOC = ""
		$049 = "MIAA"
		$ITYPE = 21
		$ITEM_STAT = "h"

;MIDDLETOWN
	Case "mdli"
		$BIB_LOC = "mdl"
		$LABELLOC = ""
		$049 = "OMMM"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "mdrs"
		$BIB_LOC = "mdl"
		$LABELLOC = ""
		$049 = "OMMM"
		$ITYPE = 0
		$ITEM_STAT = "-"
	Case "mdr"
		$BIB_LOC = "multi"
		$BIB_LOC_L1 = "mdl"
		$BIB_LOC_L2 = "ref"
		$LABELLOC = "Ref"
		$049 = "OMMM"
		$ITYPE = 21
		$ITEM_STAT = "o"
		$REF = 1
	Case "mdim"
		$BIB_LOC = "mdl"
		$LABELLOC = "IMC"
		$049 = "OMMM"
		$ITYPE = 50
		$ITEM_STAT = "-"
	Case "mdju"
		$BIB_LOC = "mdl"
		$LABELLOC = "Juv"
		$049 = "OMMM"
		$ITYPE = 0
		$ITEM_STAT = "-"
EndSwitch

;store variables for future script use
_StoreVar("$T25")
_StoreVar("$BIB_LOC")
_StoreVar("$LABELLOC")
_StoreVar("$ITYPE")
_StoreVar("$ITEM_STAT")
_StoreVar("$BIB_LOC_L1")
_StoreVar("$BIB_LOC_L2")
_StoreVar("$REF")
_StoreVar ("$049")
