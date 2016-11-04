hl7segs := ["MSH","PID","PV1","IN1","GT1","ORC","OBR","DG1","OBX","NTE"]
hl7 := Object()
;~ hl7.MSH :=	["enc"								; encoding chars
			;~ ,"sendApp"							; sending application
			;~ ,"sendFac"							; sending facility
			;~ ,"recApp"							; receiving application
			;~ ,"recFac"							; receiving facility
			;~ ,"date"								; date/time YYYYMMDDHHMMSS
			;~ ,"sec"								; security (not supported)
			;~ ,"type"								; msg type = "ORM^O01"
			;~ ,"ID"								; msg unique ID
			;~ ,"proc"								; processing: P=production, T=test
			;~ ,"ver"]								; HL7 ver = "2.3"

;~ hl7.PID :=	["num"				; sequence num starting at 1
			;~ ,"intID"			; internal ID
			;~ ,"MRN"				; MRN
			;~ ,"altID"			; patient account num
			;~ ,"name"				; Last 60^First 60^MI 60
			;~ ,"mothers"			; mother's maiden name (not supported)
			;~ ,"dob"				; YYYYMMDD
			;~ ,"sex"				; Male Female Unknown
			;~ ,"alias"			; not supported
			;~ ,"race"				; race
			;~ ,"addr"				; Addr1 60^Addr2 50^City 25^State 10^Zip 12
			;~ ,"country"			; not supported
			;~ ,"phHome"			; NNNNNNNNNN
			;~ ,"phBus"			; NNNNNNNNNN
			;~ ,"lang"				; not supported
			;~ ,"marital"			; Single Married Divorced Widowed Legally Separated Unknown Partner
			;~ ,"religion"			; not supported
			;~ ,"18"				; null
			;~ ,"SSN"]				; NNNNNNNNN

;~ hl7.PV1 :=	["num"				; serial num starting with 1
			;~ ,"class"			; not supported
			;~ ,"loc"				; not supported
			;~ ,"type"				; not supported
			;~ ,"05"				; not supported
			;~ ,"06"				; not supported
			;~ ,"attg"				; Provider Code 20^LastName 60^FirstName 60^MI 60
			;~ ,"ref"				; Provider Code 20^LastName 60^FirstName 60^MI 60
			;~ ,"09"
			;~ ,"10"
			;~ ,"11"
			;~ ,"12"
			;~ ,"13"
			;~ ,"14"
			;~ ,"15"
			;~ ,"16"
			;~ ,"17"
			;~ ,"18"
			;~ ,"req"]				; requisition number

;~ hl7.IN1 :=	["num"				; serial num starting with 1
			;~ ,"insPlan"			; not supported
			;~ ,"insCode"			; insurance company ID lab bill mnemonic (carrier code)
			;~ ,"insName"			; insurance company name
			;~ ,"insAddr"			; Addr1 40^Addr2 50^City 15^State 15^Zip 10
			;~ ,"insContact"		; not supported
			;~ ,"insPhone"			; NNNNNNNNNN
			;~ ,"groupNum"			; group number
			;~ ,"groupName"		; not supported
			;~ ,"10"				; not supported
			;~ ,"11"
			;~ ,"12"
			;~ ,"13"
			;~ ,"14"
			;~ ,"type"				; plan type
			;~ ,"name"				; name of insured Last 60^First 60^MI 60
			;~ ,"rel"				; insured's relationship to patient
			;~ ,"dob"				; insured's dob YYYYMMDD
			;~ ,"addr"]			; insured's address Addr1 60^Addr2 50^City 25^State 10^Zip 12

;~ hl7.GT1 :=	["num"				; sequence num starting with 1
			;~ ,"02"
			;~ ,"name"				; guarantor name Last 60^First 60^MI 60
			;~ ,"spouse"			; not supported
			;~ ,"addr"				; Addr1 60^Addr2 50^City 25^State 10^Zip 12
			;~ ,"phone"			; NNNNNNNNNN
			;~ ,"07"
			;~ ,"08"
			;~ ,"sex"				; see above
			;~ ,"10"
			;~ ,"rel"				; see above
			;~ ,"SSN"]				; NNNNNNNNN

;~ hl7.ORC :=	["ctrl"				; to identify new orders
			;~ ,"place"			; place order number
			;~ ,"03"
			;~ ,"04"
			;~ ,"05"
			;~ ,"06"
			;~ ,"07"
			;~ ,"08"
			;~ ,"date"				; date/time of transaction YYYYMMDDHHMMSS
			;~ ,"10"
			;~ ,"11"
			;~ ,"prov"				; ordering provider Provider code 20^LastName 60^FirstName 60^MI 60
			;~ ,"13"]

;~ hl7.OBR :=	["num"				; serial num starting with 1
			;~ ,"req"				; requisition number
			;~ ,"obs"				; observation battery identifier Code 20^TestName 255
			;~ ,"04"
			;~ ,"05"
			;~ ,"06"
			;~ ,"date"				; collection date/time YYYYMMDDHHMMSS
			;~ ,"08"
			;~ ,"09"
			;~ ,"10"
			;~ ,"11"
			;~ ,"12"
			;~ ,"13"
			;~ ,"14"
			;~ ,"source"			; specimen source Specimen Source 30^Specimen Description 255
			;~ ,"prov"				; ordering provider Provider Code 20^LastName 60^FirstName 60^MI 60
			;~ ,"17"
			;~ ,"altID"			; customer specific ID
			;~ ,"fasting"			; 0=non-fasting, 1=fasting
			;~ ,"20"
			;~ ,"21"
			;~ ,"22"
			;~ ,"23"
			;~ ,"24"
			;~ ,"25"
			;~ ,"26"
			;~ ,"stat"				; 0=regular, 1=STAT
			;~ ,"cc"]				; CC List ~ separated Provider Code 20^LastName 60^FirstName 60^MI 60

;~ hl7.OBX :=	["num"				; serial num starting with 1
			;~ ,"type"				; standard text "ST"
			;~ ,"obs"				; observation identifier Question code 20^Question 50
			;~ ,"04"
			;~ ,"ans"]				; Answer 50^Answer Code (Optional and for Drop down only) 20

;~ hl7.DG1 :=	["num"				; serial num starting with 1
			;~ ,"02"
			;~ ,"code"				; diagnosis code 
			;~ ,"dx"]				; diagnosis description

;~ hl7.NTE :=	["num"				; serial num starting with 1
			;~ ,"02"
			;~ ,"note"]			; comment

	hl7.MSH := []
	hl7.PID := []
	hl7.PV1 := []
	hl7.IN1 := []
	hl7.GT1 := []
	hl7.DG1 := []
	hl7.NTE := []
	hl7.ORC := []
	hl7.OBR := []
	hl7.OBX := []
	
hl7.MSH[2] := "sdf"
;MsgBox % hl7.msh.2
MsgBox % IsObject(hl7.nte)
ExitApp

FileRead, txt, samples\hl7test.txt

loop, parse, txt, `n, `r
{
	seg := A_LoopField
	StringSplit, fld, seg, |
	loop, % fld0
	{
		n := A_Index
		if (n=1) 
			continue
		m := fld%n%
		i := hl7[fld1][n-1]
		j := hl7[fld1][i]
		MsgBox,, % n "-" hl7[fld1][n-1], % m
	}
	segName := fld1
	MsgBox,, % segname, % ObjHasValue(hl7segs, segName)
}


ExitApp

ObjHasValue(aObj, aValue, rx:="") {
; modified from http://www.autohotkey.com/board/topic/84006-ahk-l-containshasvalue-method/	
    for key, val in aObj
		if (rx) {
			if (med) {													; if a med regex, preface with "i)" to make case insensitive search
				val := "i)" val
			}
			if (aValue ~= val) {
				return, key, Errorlevel := 0
			}
		} else {
			if (val = aValue) {
				return, key, ErrorLevel := 0
			}
		}
    return, false, errorlevel := 1
}

