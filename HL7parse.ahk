initHL7()

FileRead, txt, samples\hl7test.txt

loop, parse, txt, `n, `r																; parse HL7 message, split on `n, ignore `r for Unix files
{
	seg := A_LoopField																	; read next Segment line
	if (seg=="") {
		continue
	}
	fld := StrSplit(seg,"|")															; split on `|` field separator into fld pseudo-array
		fld.0 := fld.length()
		segName := fld.1																; first array element should be NAME
	if !IsObject(hl7[segName]) {
		MsgBox,,% A_Index, % seg "-" segName "`nBAD SEGMENT NAME"
		continue																		; skip if segment name not allowed
	}
	out := hl7sep(fld)
	if (segName="OBX") {																; need to special process OBX[3], test result strings
		if (out.Filename) {																; file follows
			MsgBox % out.Filename
		} else {
			lab := out.resName															; label is actually the component
			res := (out.resValue) 
				? out.resValue . (out.resUnits ? " " out.resUnits : "")					; [5] value and [6] units
				: ""
			MsgBox % lab ": " res
		}
	}
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

hl7sep(fld) {
	global hl7
	res := Object()
	segName := fld.1
	segMap := hl7[segName]
	Loop, % fld.0
	{
		i := A_Index																	; step through each of the fld[] strings
		str := fld[i]
		strMap := segMap[i-1]															; get hl7 substring that maps to this 
		if (strMap=="") {																; no matching string map
			if !(str=="") {																; but a value
				res[i-1] := str															; create a [0] marker
			}
			continue
		}
		map := StrSplit(strMap,"^")														; array of substring map
		val := StrSplit(str,"^")														; array of subelements
		loop, % map.length()
		{
			j := A_Index
			res[map[j]] := val[j]														; add each subelement
		}
	}
	return res
}
