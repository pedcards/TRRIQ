initHL7()
fldVal := []																		; ///	Use this for testing

FileRead, txt, samples\hl7test.txt

loop, parse, txt, `n, `r																; parse HL7 message, split on `n, ignore `r for Unix files
{
	seg := A_LoopField																	; read next Segment line
	if (seg=="") {
		continue
	}
	out := hl7sep(seg)
}

ExitApp

initHL7() {
	global hl7
	hl7 := Object()
	IniRead, s0, hl7.ini																	; s0 = Section headers
	loop, parse, s0, `n, `r																	; parse s0
	{
		i := A_LoopField
		hl7[i] := []																		; create array for each header
		IniRead, s1, hl7.ini, % i															; s1 = individual header
		loop, parse, s1, `n, `r																; parse s1
		{
			j := A_LoopField
			arr := strSplit(j,"=",,2)														; split into arr.1 and arr.2
			hl7[i][arr.1] := arr.2															; set hl7.OBX.2 = "Obs Type"
		}
	}
	return
}

hl7sep(seg) {
	global hl7, fldVal
	res := Object()
	fld := StrSplit(seg,"|")															; split on `|` field separator into fld array
	segName := fld.1																	; first array element should be NAME
	segMap := hl7[segName]
	if !IsObject(hl7[segName]) {														; no matching hl7 map?
		MsgBox,,% A_Index, % seg "-" segName "`nBAD SEGMENT NAME"
		return error																	; fail if segment name not allowed
	}
	Loop, % fld.length()																; step through each of the fld[] strings
	{
		i := A_Index
		str := fld[i]																	; each segment field
		strMap := segMap[i-1]															; get hl7 substring that maps to this 
		if (strMap=="") {																; no matching string map
			if !(str=="") {																; but a value
				res[i-1] := str															; create a [n] marker
			}
			continue
		}
		map := StrSplit(strMap,"^")														; array of substring map
		val := StrSplit(str,"^")														; array of subelements
		loop, % map.length()
		{
			j := A_Index
			res[map[j]] := val[j]														; add each subelement
			if !(fldVal[map[j]]) {
				fldVal[map[j]] := val[j]
			}
		}
	}
	if (segName="OBX") {																	; need to special process OBX[3], test result strings
		if !(res.Filename=="") {																; file follows
			fldVal.Filename := res.Filename
			;~ nBytes := Base64Dec( out.resValue, Bin )
			;~ File := FileOpen( out.Filename, "w")
			;~ File.RawWrite(Bin, nBytes)
			;~ File.Close()
		} else {
			label := res.resCode															; label is actually the component
			result := (res.resValue) 
				? res.resValue . (res.resUnits ? " " res.resUnits : "")					; [5] value and [6] units
				: ""
			fldVal[label] := result
		}
	}
	
	return res
}

Base64Dec( ByRef B64, ByRef Bin ) {  ; By SKAN / 18-Aug-2017
; from https://autohotkey.com/boards/viewtopic.php?t=35964
Local Rqd := 0, BLen := StrLen(B64)                 ; CRYPT_STRING_BASE64 := 0x1
  DllCall( "Crypt32.dll\CryptStringToBinary", "Str",B64, "UInt",BLen, "UInt",0x1
         , "UInt",0, "UIntP",Rqd, "Int",0, "Int",0 )
  VarSetCapacity( Bin, 128 ), VarSetCapacity( Bin, 0 ),  VarSetCapacity( Bin, Rqd, 0 )
  DllCall( "Crypt32.dll\CryptStringToBinary", "Str",B64, "UInt",BLen, "UInt",0x1
         , "Ptr",&Bin, "UIntP",Rqd, "Int",0, "Int",0 )
Return Rqd
}

