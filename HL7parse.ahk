hl7 := Object()
IniRead, s0, hl7.ini
loop, parse, s0, `n, `r
{
	i := A_LoopField
	hl7[i] := []
	IniRead, s1, hl7.ini, % i
	loop, parse, s1, `n, `r
	{
		j := A_LoopField
		arr := strSplit(j,"=",,2)
		hl7[i][arr[1]] := arr[2]
	}
}
	
FileRead, txt, samples\hl7test.txt

loop, parse, txt, `n, `r																; parse HL7 message, split on `n, ignore `r for Unix files
{
	seg := A_LoopField																	; read next Segment line
	StringSplit, fld, seg, |															; split on `|` field separator into fld pseudo-array
		segNum := fld0																	; number of elements from StringSplit
		segName := fld1																	; first array element should be NAME
	if !IsObject(hl7[segName]) {
		MsgBox,,% segName, BAD SEGMENT NAME
		continue																		; skip if segment name not allowed
	}
	loop, % segNum
	{
		hl7[segName][A_Index-1] := fld%A_Index%											; start counting at 0
	}
	if (segName="MSH") {
		if !(hl7.MSH.8="ORU^R01") {
			MsgBox % hl7.MSH.8 "`nWrong message type"
			break
		}
	}
	if (segName="OBX") {																; need to process each OBX in turn during loop
		hl7sep("OBX",3)
	}
}

ExitApp

hl7sep(seg,fld) {
	global hl7
	str := hl7[seg][fld]																; Field string to separate
	map := hl7[seg].map[fld]															; Equivalent map to separate
	StringSplit, cmp, str, `^															; Split string into components
	StringSplit, val, map, `^															; Split map into text values
	
	if (seg="OBX" && fld=3) {															; need to special process OBX[3], test result strings
		lab := cmp1																		; label is actually the component
		res := hl7.OBX.5 " " hl7.OBX.6													; [5] value and [6] units
		MsgBox % lab ": " res
		return
	}
	loop, % val0																		; perform as long as map string exists
	{
		lab := val%A_Index%																; generate label and value pairs
		res := cmp%A_Index%																; for each component, single or multiple
		
		MsgBox % lab "`n" res
	}
}
