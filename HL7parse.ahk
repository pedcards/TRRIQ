hl7segs := ["MSH","PID","PV1","IN1","GT1","ORC","OBR","DG1","OBX","NTE"]
hl7 := Object()
hl7.MSH := []
/*	The Message Header (MSH) segment defines the intent, source, destination, and some specifics of the syntax of a message.
	FIELD NAME						SEG. NAME	COMMENT											LEN		R/O/N
	Segment Type ID					MSH.00		MSH												3		R
	Encoding Characters				MSH.01		^~|&											4		R
	Sending Application				MSH.02		Sending Application								50		R
	Sending Facility				MSH.03		Sending Facility								50		R
	Receiving Application			MSH.04		Receiving Application							50		R
	Receiving Facility				MSH.05		Receiving Facility								50		R
	Message Date & Time				MSH.06		Format: YYYYMMDDHHMMSS							12		R
	Security						MSH.07		Not Supported											N
	Message Type					MSH.08		ORM^O01													R
	Message Control ID				MSH.09		A number that uniquely identifies the message			R
	Processing ID					MSH.10		P-Production; T-Test							1		R
	Version ID						MSH.11		2.3												8		R
*/
hl7.PID := []
/*	The Patient Identifier (PID) segment is used by all application as the primary means of communication patient identification information. This segment contains permanent patient identifying and demographic information that, for the most part, is not likely to change frequently.
	Segment Type ID					PID.00		PID												3		R
	Sequence Number					PID.01		Serial Number starting from 1					4		R
	Patient ID (Internal ID)		PID.02		Patient ID (Internal ID)						20		O
	Patient ID (External ID)		PID.03		Patient MRN										20		O
	Alternate Patient ID			PID.04		Client Patient Account Number					20		O
	Patient Name					PID.05		Last 60^First 60^MI 60							180		R
	Mother’s Maiden Name			PID.06		Not Supported											N
	Date of Birth					PID.07		YYYYMMDD										8		R
	Sex								PID.08		Male Female Unknown										R
	Patient Alias					PID.09		Not Supported											N
	Race							PID.10		See Appendix A.2										O
	Patient Address					PID.11		Addr1 60^Addr2 50^City 25^State 10^Zip 12		161		O
	County Code						PID.12		Not Supported											N
	Phone # - Home					PID.13		NNNNNNNNNN										10		O
	Phone # - Business				PID.14		NNNNNNNNNN										10		O
	Primary Language				PID.15		Not Supported											N
	Marital Status					PID.16		See Appendix A.3										O
	Religion						PID.17		Not Supported											N
	SSN –Patient					PID.19		NNNNNNNNN										9		O
*/
hl7.PV1 := []
/*	The PV1 segment is used to communicate information on a visit-specific basis
	Segment Type ID					PV1.00		PV1												3		R
	Sequence Number					PV1.01		Serial Number starting from 1					4		R
	Attending Doctor				PV1.07		Provider Code 20^LastName 60^FirstName 60^MI 60	200		O
	Referring Provider				PV1.08		Provider Code 20^LastName 60^FirstName 60^MI 60	200		O
	Visit Number					PV1.19		Requisition Number								11		O
*/
hl7.IN1 := []
/*	The Insurance (IN1) segment contains insurance policy coverage information necessary to produce properly pro-rated and patient and insurance bills. This segment is applicable only to the outbound order for insurance billing.
	Segment Type ID					IN1.00		IN1												3		R
	Sequence Number					IN1.01		Serial Number starting from 1					4		R
	Insurance Plan ID				IN1.02		Not Supported											N
	Insurance Company ID			IN1.03		Lab Bill mnemonic (carrier code)				20		O
	Insurance Company Name			IN1.04		Insurance Company Name							40		R
	Insurance Company Address		IN1.05		Addr1 40^Addr2 50^City 15^State 15^Zip 10		130		R
	Insurance Co. Contact Person	IN1.06		Not Supported											N
	Insurance Co. Phone Number		IN1.07		NNNNNNNNNN										10		O
	Group Number					IN1.08		Group Number									30		O
	Plan Type						IN1.15		Plan Type										8		O
	Name of Insured					IN1.16		Last 60^First 60^MI 60							180		O
	Insured’s Rel to Patient		IN1.17		See Appendix A.3										R
	Insured’s Date of Birth			IN1.18		YYYYMMDD										8		R
	Insured’s Address				IN1.19		Addr1 60^Addr2 50^City 25^State 10^Zip 12		157		R
	Policy Number					IN1.36		Same as subscriber number						20		O
	Bill Type						IN1.47		P-Patient C-Client T-Third Party Bill			1		O
*/
hl7.GT1 := []
/*	The Guarantor (GT1) segment contains guarantor (for example, the person or the organization with financial responsibility for payment of a patient account) data for patient and insurance billing applications. This segment is applicable only to the outbound order for patient and insurance billing.
	Segment Type ID					GT1.00		GT1												3		R
	Sequence Number					GT1.01		Serial Number starting from 1					4		R
	Guarantor Number				GT1.02		Not Supported											N
	Guarantor Name					GT1.03		Last 60^First 60^MI 60							180		R
	Guarantor Spouse Name			GT1.04		Not Supported											N
	Guarantor Address				GT1.05		Addr1 60^Addr2 50^City 25^State 10^Zip 12		161		R
	Guarantor Ph Num-Home			GT1.06		NNNNNNNNNN										10		R
	Guarantor Date of Birth			GT1.07		Not Supported											N
	Guarantor Sex					GT1.09		See Appendix A.1										R
	Guarantor Type					GT1.10		Not Supported											N
	Guarantor Relationship			GT1.11		See Appendix A.3								2		R
	Guarantor SSN					GT1.12		NNNNNNNNN										9		O
	Guarantor Date – Begin			GT1.13		Not Supported											N
	Guarantor Date – End			GT1.14		Not Supported											N
*/
hl7.ORC := []
/*	The Common Order (ORC) segment is used to transmit fields that are common to all orders (all types of service that are requested). The ORC segment is required in the ORM message.
	Segment Type ID					ORC.00		ORC												3		R
	Order Control					ORC.01		To identify new orders							4		R
	Place Order Number				ORC.02		Requisition Number								11		R
	Date/Time of Transaction		ORC.09		YYYYMMDDHHMMSS									14		R
	Ordering Provider				ORC.12		Provider code 20^LastName 60^FirstName 60^MI 60	200		O
	Enterer’s Location				ORC.13		Not Supported											N
*/
hl7.OBR := []
/*	At least on OBR segment is transmitted for each Order Code associated with any PID segment. This segment is mandatory in ORM messages.
	Segment Type ID					OBR.00		OBR												3		R
	Sequence No						OBR.01		Serial Number starting from 1					4		R
	Placer Order Number				OBR.02		Requisition Number								25		R
	Filler Order Number				OBR.03		Not Supported											N
	Obs Battery Identifier			OBR.04		Code 20^TestName 255							275		R
	Obs Collection Date/Time #		OBR.07		YYYYMMDDHHMMSS									14		R
	Specimen Source					OBR.15		Specimen Source 30^Specimen Description 255		285		O
	Ordering Provider				OBR.16		Provider Code 20^LastName 60^FirstName 60^MI 60	200		R
	Alternate Specimen ID			OBR.18		Customer Specific ID							40		O
	Fasting							OBR.19		0–Non Fasting, 1–Fasting						1		O
	Priority/Stat					OBR.27		0–Regular, 1–Stat								1		O
	CC copies to					OBR.28		CC List ~ separated
												Provider Code 20^LastName 60^FirstName 60^MI 60	200+	O
*/
hl7.OBX := []
/*	This segment is optional. Ask at Order Entry Questions in the order are typically captured as OBX segments.
	Segment Type ID					OBX.00		OBX												3		R
	Sequence Number					OBX.01		Serial Number starting from 1					4		R
	Type Value						OBX.02		ST – Standard Text								2		R
	Observation Identifier			OBX.03		Question code 20^Question 50					70		R
	Answer							OBX.05		Answer 50^Answer Code (Optional and for Drop down only) 20	70	R
*/
hl7.DG1 := []
/*	The Diagnosis (DG1) segment contains patient diagnosis information, and is present on ORM messages if associated with the test. It allows identification of multiple diagnosis segments grouped beneath a single OBR segment.
	Segment Type ID					DG1.00		DG1												3		R
	Sequence Number					DG1.01		Serial Number starting from 1					4		R
	Diagnosis Coding Method			DG1.02		Not Supported											N
	Diagnosis Code					DG1.03		Diagnosis Code									60		R
	Diagnosis Name					DG1.04		Diagnosis description							60		R
*/
hl7.NTE := []
/*	The Notes and Comments (NTE) segment contains notes and comments for ORM messages, and is optional.
	Segment Type ID					NTE.00		NTE												3		R
	Sequence Number					NTE.01		Serial Number starting from 1					4		R
	Source and Comment				NTE.02		Not Supported											N
	Comment							NTE.03		Comment											255		R
*/
	
;~ hl7.MSH[2] := "sdf"
;~ fld := "MSH"
;~ MsgBox % hl7[fld].2
;~ ;MsgBox % IsObject(hl7[fld])
;~ ExitApp

FileRead, txt, samples\hl7test.txt

loop, parse, txt, `n, `r																; parse HL7 message, split on `n, ignore `r for Unix files
{
	seg := A_LoopField																	; read next Segment line
	StringSplit, fld, seg, |															; split on `|` field separator into fld pseudo-array
		segNum := fld0																	; number of elements from StringSplit
		segNam := fld1																	; first array element should be NAME
	if !IsObject(hl7[segNam]) {
		MsgBox BAD SEGMENT NAME
		continue																		; skip if segment name not allowed
	}
	loop, % segNum
	{
		n := A_Index																	; start counting at 0
		hl7[segNam][n-1] := fld%n%
		;~ MsgBox,, % n-1, % hl7[segNam][n-1]
	}
}

MsgBox % hl7.msh.1

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

