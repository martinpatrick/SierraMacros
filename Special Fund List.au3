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

 Name of script: Special Fund List

 Script Function:
	This script decides whether or not a bib record needs notes due to being
	purchased with donated/endowed funds, and then stores the data to the registry
	for retrieval by other scripts in the set.

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
Dim $FUND, $FUND2
Dim $SF_NAME ;used for order record notes
Dim $SF_590 ;bib record field
;Dim $SF_590_2 ;second bib record field if needed
Dim $SF_7XX ;bib record fields, can contain multiple 79X entries
Dim $SF_FIB ;prompts the user to fill in the blank during the D button function

;################################ MAIN ROUTINE #################################
;load fund variable from other scripts
$FUND = _LoadVar("$FUND2")
;MsgBox(0, "hey!", "checking the fund code! which is " & $FUND)


	If $FUND = "4b " Then; the macro grabs the first three characters of the fund code (easiest method) so two digit fund codes need to have an extra space for the match
		$SF_NAME = "Baer"
		$SF_590 = "purchased with monies from the Paul W. & Elsa Thoma Baer Fund"
		$SF_7XX = "b7901 Baer, Elsa Thoma,|edonor" & @CR & "b7901 Baer, Paul W.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4bi" Then
		$SF_NAME = "Bierly"
		$SF_590 = "donated in honor of Dr. Willis (Andy) W. Wertz, MU 1932, Dept. of Architecture, 1937-1973, by RIchard Bierly, MU 1954."
		$SF_7XX = "b7901 Bierly, Richard,|edonor" & @CR & "b7901 Wertz, Willis W.,|ehonoree"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4c " Then
		$SF_NAME = "Wertz"
		$SF_590 = "purchased with monies from the Willis W. Wertz Library (Architecture) Acquisitions Fund"
		$SF_7XX = "b7901 Wertz, Willis W.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4dr" Then
		$SF_NAME = "Drake"
		$SF_590 = "purchased with monies from the  the Vivian L. Drake Fund"
		$SF_7XX = "b7901 Drake, Vivian L.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4du" Then
		$SF_NAME = "Duncan"
		$SF_590 = "purchased with monies from the Richard & Dorothy Duncan Fund"
		$SF_7XX = "b7901 Duncan, Richard,|edonor" & @CR & "b7901 Duncan, Dorothy,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4e " Then
		$SF_NAME = "Class of 1920"
		$SF_590 = "purchased with monies from the Class of 1920 University Libraries Enrichment Fund"
		$SF_7XX = "b7912 Class of 1920,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4eb" Then
		$SF_NAME = "Elec.Bus.Res."
		$SF_590 = "purchased with monies from the University Libraries Electronic Business Resources Fund"
		$SF_7XX = "b7912 University Libraries Electronic Business Resources,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4er" Then
		$SF_NAME = "Erwin"
		$SF_590 = "purchased with monies from the David & Susan Erwin Library Acquisitions Fund"
		$SF_7XX = "b7901 Erwin, David,|edonor" & @CR & "b7901 Erwin, Susan,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4f " Then
		$SF_NAME = "Mozingo"
		$SF_590 = "purchased with monies from the Todd Mozingo Memorial Fund"
		$SF_7XX = "b7901 Mozingo, Todd,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4ft" Then ;clipped because it's looking for three characters only
		$SF_NAME = "Foote"
		$SF_590 = "purchased in memory of Diane Altman Ball '54"
		$SF_7XX = "b7901 Ball, Diane Altman,|ehonoree" & @CR & "b7901 Foote, Janet,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4gb" Then
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "b7901 Hovorka, Sophie Nickel,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4gg" Then
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "b7901 Hovorka, Sophie Nickel,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4gp" Then
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "b7901 Hovorka, Sophie Nickel,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4h " Then
		$SF_NAME = "G.Hill R. R."
		$SF_590 = "purchased with monies from the Gretchen Hill Children’s Literature Reading Area Fund"
		$SF_7XX = "b7901 Hill, Gretchen,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4ia" Then
		$SF_NAME = "Iandoli"
		$SF_590 = "purchased with monies from the Marie Iandoli Library Acquisitions Fund"
		$SF_7XX = "b7901 Iandoli, Marie,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4j " Then
		$SF_NAME = "Wesolowski"
		$SF_590 = "purchased with monies from the John Wesolowski Memorial Libraries Endowment"
		$SF_7XX = "b7901 Wesolowski, John,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4k " Then
		$SF_NAME = "I. Williams"
		$SF_590 = "This item was acquired through funding from the William D. Mulliken Library Acquisitions Fund"
		$SF_7XX = "b7901 Williams, Isabella Riggs,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4l " Then
		$SF_NAME = "Mulliken"
		$SF_590 = "purchased with monies from the Isabella Riggs Williams Library Fund"
		$SF_7XX = "b7901 Mulliken, William D.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4ky" Then
		$SF_NAME = "Kirby"
		$SF_590 = "purchased in honor of Dr. Jack T. Kirby, Emeritus, Professor, History (1965-2002) by FILL IN THIS BLANK"
		$SF_7XX = "b7901 Kirby, Jack T.,|ehonoree" & @CR & "b7901 FILL IN THE BLANK"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4lf" Then
		$SF_NAME = "Lyons-Felstead"
		$SF_590 = "purchased in memory of Anna Mae Duvall Lyons '32"
		$SF_7XX = "b7901 Lyons, Anna Mae Duvall,|ehonoree" & @CR & "b7901 Lyons-Felstead, Joan,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4m " Then
		$SF_NAME = "Sinclair"
		$SF_590 = "purchased with monies from the Robert B. Sinclair Library Collections Fund"
		$SF_7XX = "b7901 Sinclair, Robert B.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4n " Then
		$SF_NAME = "Schlachter"
		$SF_590 = "purchased with monies from the Karl & Roberta Schlachter Library Fund"
		$SF_7XX = "b7901 Schlachter, Karl,|edonor" & @CR & "b7901 Schlachter, Roberta,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4naw" Then
		$SF_NAME = "NAWPA"
		$SF_590 = "purchased with monies from the Native American Women Playwrights Archive Fund"
		$SF_7XX = "b7912 Native American Women Playwrights Archive,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4o " Then
		$SF_NAME = "Havighurst"
		$SF_590 = "purchased with monies from the Walter E. Havighurst Special Collections Library Fund"
		$SF_7XX = "b7901 Havighurst, Walter R.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4p " Then
		$SF_NAME = "Peterson/Defoe"
		$SF_590 = "purchased with monies from the Spiro Peterson Center for Defoe Studies"
		$SF_7XX = "b7901 Peterson, Spiro,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4pr" Then
		$SF_NAME = "Pratt Amer. Lit."
		$SF_590 = "purchased with monies from the William C. Pratt American Literature Acquisitions Fund"
		$SF_7XX = "b7901 Pratt, William C.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4q " Then
		$SF_NAME = "NEH"
		$SF_590 = "purchased with monies from the NEH Challenge Fund"
		$SF_7XX = "b7912 NEH Challenge,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4r " Then
		$SF_NAME = "Alumni"
		$SF_590 = "purchased with monies donated by "
		$SF_7XX = "b7911 "
		$SF_FIB = 1
		_StoreVar("$SF_FIB")
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4s " Then
		$SF_NAME = "Covington-Williams"
		$SF_590 = "purchased with monies from the Covington-Williams Families Fund"
		$SF_7XX = "b7903 Covington-Williams Families,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4sc" Then ; should be 4sch, but 4sc doesn't bump into anything and since we're only dealing in the first three chars.
		$SF_NAME = "Shafer"
		$SF_590 = "purchased with monies from the Alice Schafer Fund"
		$SF_7XX = "b7901 Schafer, Alice,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4sm" Then
		$SF_NAME = "Swenson-Mount"
		$SF_590 = "given in memory of Edward & Gladys Swenson and Robert L. Mount Children’s Literature Fund"
		$SF_7XX = "b7901 Swenson, Edward,|ehonoree" & @CR & "b7901 Swenson, Gladys,|ehonoree" & @CR & "b7901 Mount, Robert L.,|ehonoree"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4t " Then
		$SF_NAME = "Brown-Lane"
		$SF_590 = "purchased with monies from the Mariyette Brown-Lane Library Endowment Fund"
		$SF_7XX = "b7901 Brown-Lane, Mariyette,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4u " Then
		$SF_NAME = "Halstead"
		$SF_590 = "purchased with monies from the Halstead Collection Fund"
		$SF_7XX = "b7903 Halstead Collection,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4v " Then
		$SF_NAME = "Richey"
		$SF_590 = "purchased with monies from the Sam W. Richey Collection of Southern Confederacy Fund"
		$SF_7XX = "b7901 Richey, Sam. W.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4wa" Then
		$SF_NAME = "Joanne Ward IMC"
		$SF_590 = "purchased with monies from the Joanne Stalzer Ward Instructional Materials Center Support Fund"
		$SF_7XX = "b7901 Ward, Joanne Stalzer,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4x " Then
		$SF_NAME = "Fletcher"
		$SF_590 = "purchased with monies from the Alice Harries Fletcher Special Collections Fund"
		$SF_7XX = "b7901 Fletcher, Alice Harries,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "4y " Then
		$SF_NAME = "Sissman"
		$SF_590 = "purchased with monies from the Ben Sissman-WWII Intelligence History Library Fund"
		$SF_7XX = "b7901 Sissman, Ben,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "54 " Then
		$SF_NAME = "Dean’s Special Fund"
		$SF_590 = "purchased with monies from the University Libraries Acquisitions Fund"
		$SF_7XX = "b7912 University Libraries Acquisitions Fund,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "wes" Then
		$SF_NAME = "Weslowski"
		$SF_590 = "purchased with monies from the John Weslowski Memorial Libraries Endowment Fund"
		$SF_7XX = "b7901 Weslowksi, John,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "1vc" Then
		$SF_NAME = "SUPRONAS"
		$SF_590 = "purchased in Memory of Ainis (Andy) Supronas with funds donated in support of Baltic Studies"
		$SF_7XX = "Supronas, Ainis (Andy),|ehonoree"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $FUND = "11h" Then
		$SF_NAME = "SUPRONAS"
		$SF_590 = "purchased in Memory of Ainis (Andy) Supronas with funds donated in support of Baltic Studies"
		$SF_7XX = "Supronas, Ainis (Andy),|ehonoree"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	Else
		$SF_NAME = 0
		_StoreVar("$SF_NAME")
	EndIf

Exit

;store variables for use in other scripts
;_StoreVar("$SF_NAME")
;_StoreVar("$SF_590")
;_StoreVar("$SF_7XX")
