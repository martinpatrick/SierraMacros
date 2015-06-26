#cs
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
#ce

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=images\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <TSCustomFunction.au3>

$var = EnvGet("TEMP")
MsgBox(4096, "Path variable is:", $var)
;~ ClipPut($var)
$var = StringTrimLeft($var, 12)
$var = StringTrimRight($var, 14)

Switch $var
	Case "alexanpk"
		$C_INI = "pk"
	Case "barbouh2"
		$C_INI = "hlb"
	Case "patricm"
		$C_INI = "mp"
	Case "spencert"
		$C_INI = "rts"
	Case "keyessl"
		$C_INI = "sk"
	Case "stepanm"
		$C_INI = "ms"
    Case "bazelejw"
		$C_INI = "jwb"
    Case "abneymd"
		$C_INI = "ma"
    Case "smithjl9"
	    $C_INI = "js"
    Case "cliftks"
	    $C_INI = "kc"
	case Else
		$C_INI = "999"
EndSwitch
msgbox(0, "ini", $C_INI)

_StoreVar("$C_INI")
