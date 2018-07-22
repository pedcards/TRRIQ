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

processHL7(fnam) {
	global hl7Dir
	FileRead, txt, % hl7Dir . fnam
	StringReplace, txt, txt, `r`n, `r														; convert `r`n to `r
	StringReplace, txt, txt, `n, `r															; convert `n to `r
	loop, parse, txt, `r, `n																; parse HL7 message, split on `r, ignore `n
	{
		seg := A_LoopField																	; read next Segment line
		if (seg=="") {
			continue
		}
		hl7line(seg)
	}
	return
}

hl7line(seg) {
/*	Interpret an hl7 message "segment" (line)
	segments are comprised of fields separated by "|" char
	field elements can contain subelements separated by "^" char
	field elements stored in res[i] object
	attempt to map each field to recognized structure for that field element
*/
	global hl7, fldVal, hl7Dir
	res := Object()
	fld := StrSplit(seg,"|")															; split on `|` field separator into fld array
	segName := fld.1																	; first array element should be NAME
	if !IsObject(hl7[segName]) {														; no matching hl7 map?
		MsgBox,,% A_Index, % seg "-" segName "`nBAD SEGMENT NAME"
		return error																	; fail if segment name not allowed
	}
	
	isOBX := (segName == "OBX")
	isResult := (fld.3 ~= "^(CE|NM|ST|TX|ED)$")
	segMap := hl7[segName]
	segPre := (isOBX) ? "" : segName "_"
	segPre .= (isResult) ? fld.3 "_" : ""
	
	Loop, % fld.length()																; step through each of the fld[] strings
	{
		i := A_Index
		if (i<=2) {																		; skip first 2 elements in OBX|2|TX
			continue
		}
		str := fld[i]																	; each segment field
		strMap := segMap[i-1]															; get hl7 substring that maps to this 
		val := StrSplit(str,"^")														; array of subelements
		
		if (strMap=="") {																; no mapped fields
			loop, % val.length()														; create strMap "^^^" based on subelements in val
			{
				strMap .= "z" i "_" A_Index "^"
			}
		}
		map := StrSplit(strMap,"^")														; array of substring map
		loop, % map.length()
		{
			j := A_Index
			if (map[j]=="") {															; skip if map value is null
				continue
			}
			x := segPre . map[j]														; res.pre_map
			
			res[x] := val[j]															; add each mapped result as subelement, res.mapped_name
			
			if !(isOBX) || (isResult) {													; non-OBX results
				fldVal[x] := val[j]														; populate all fldVal.mapped_name
			}
		}
	}
	if (isOBX) {																		; need to special process OBX[], test result strings
		if !(res.ED_Filename == "") {
			fldVal.Filename := res.ED_Filename												; file follows
			nBytes := Base64Dec( res.ED_resValue, Bin )
			File := FileOpen( hl7Dir . res.ED_Filename, "w")
			File.RawWrite(Bin, nBytes)
			File.Close()
		} else {
			label := res.resCode														; result value
			result := strQ(res.resValue, "###" strQ(res.resUnits," ###"))
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

