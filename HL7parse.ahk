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
	Message Type					MSH.08		ORU^R01													R
	Message Control ID				MSH.09		A number that uniquely identifies the message			R
	Processing ID					MSH.10		P-Production; T-Test							1		R
	Version ID						MSH.11		2.3												8		R
*/
hl7.PID := []
hl7.PID.map := {2:"MRN",5:"Last Name^First Name^MI",7:"DOB",8:"Sex"}
/*	The Patient Identifier (PID) segment is used by all application as the primary means of communication patient identification information. This segment contains permanent patient identifying and demographic information that, for the most part, is not likely to change frequently.
	Segment Type ID					PID.00		PID												3		R
	Sequence Number					PID.01		Serial Number starting from 1					4		R
	Patient ID (Internal ID)		PID.02		Patient ID (Internal ID)						20		O
	Patient ID (External ID)		PID.03		Patient MRN										20		O
	Alternate Patient ID			PID.04		Client Patient Account Number					20		O
	Patient Name					PID.05		Last 60^First 60^MI 60							180		R
	Mother�s Maiden Name			PID.06		Not Supported											N
	Date of Birth					PID.07		YYYYMMDD										8		R
	Sex								PID.08		M-Male F-Female U-Unknown								R
	Patient Address					PID.11		Addr1^Addr2^City^State^Zip						106		O
	County Code						PID.12		Not Supported											N
	Phone # - Home					PID.13		NNNNNNNNNN										10		O
	Patient Account Number			PID.18		Client Patient Account Number					20		O
	SSN �Patient					PID.19		NNNNNNNNN										9		O
*/
hl7.PV1 := []
hl7.PV1.map := {7:"Attg code^Attg NameL^Attg NameF^Attg NameMI",8:"Ref code^Ref NameL^Ref NameF^Ref NameMI"}
/*	The PV1 segment is used to communicate information on a visit-specific basis and is not a required segment for the ORU Message.
	Segment Type ID					PV1.00		PV1												3		R
	Sequence Number					PV1.01		Serial Number starting from 1					4		R
	Assigned Patient Location		PV1.03		Account Number											O
	Attending Doctor				PV1.07		Provider Code 20^LastName 60^FirstName 60^MI 60	200		O
	Referring Provider				PV1.08		Provider Code 20^LastName 60^FirstName 60^MI 60	200		O
	Visit Number					PV1.19		Customer Specific Accessioning					11		O
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
	Insured�s Rel to Patient		IN1.17		See Appendix A.3										R
	Insured�s Date of Birth			IN1.18		YYYYMMDD										8		R
	Insured�s Address				IN1.19		Addr1 60^Addr2 50^City 25^State 10^Zip 12		157		R
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
	Guarantor Date � Begin			GT1.13		Not Supported											N
	Guarantor Date � End			GT1.14		Not Supported											N
*/
hl7.DG1 := []
/*	The Diagnosis (DG1) segment contains patient diagnosis information, and is present on ORM messages if associated with the test. It allows identification of multiple diagnosis segments grouped beneath a single OBR segment.
	Not used in RESULTS section.
*/
hl7.ORC := []
hl7.ORC.map := {9:"Date",12:"Ref code^Ref NameL^Ref NameF^Ref NameMI"}
/*	The Common Order (ORC) segment is used to transmit fields that are common to all orders (all types of service that are requested). The ORC segment is required in the ORM message.
	Segment Type ID					ORC.00		ORC												3		R
	Order Control					ORC.01		To identify new orders							4		N
	Place Order Number				ORC.02		Requisition Number								11		R
	Date/Time of Transaction		ORC.09		YYYYMMDDHHMMSS									14		O
	Ordering Provider				ORC.12		Provider code 20^LastName 60^FirstName 60^MI 60	200		R
	Enterer�s Location				ORC.13		Not Supported											N
*/
hl7.OBR := []
hl7.OBR.map := {4:"Test code^Test name",7:"date",25:"status"}
/*	At least on OBR segment is transmitted for each Order Code associated with any PID segment. This segment is mandatory in ORM messages.
	Segment Type ID					OBR.00		OBR												3		R
	Sequence No						OBR.01		Serial Number starting from 1					4		R
	Placer Order Number				OBR.02		Requisition Number								25		R
	Filler Order Number				OBR.03		Not Supported											N
	Obs Battery Identifier			OBR.04		Code 20^TestName 255							275		R
	Obs Collection Date/Time #		OBR.07		YYYYMMDDHHMMSS									26		R
	Specimen Source					OBR.15		Specimen Source 30^Specimen Description 255		285		O
	Ordering Provider				OBR.16		Provider Code 20^LastName 60^FirstName 60^MI 60	200		R
	Alternate Specimen ID			OBR.18		Alternate Customer Specific ID					20		O
	Performing Lab					OBR.23		labName^labAddr1^labAddr2^labAddrCity^			235		O
												labAddrState^labAddrZip^labPhone^
												labDirTitle^labDirName
	Result Status					OBR.25		P�Preliminary F�Final C�Corrected				1		O
	Courtesy Copies to				OBR.28		Provider Code^LastName ^FirstName				140		O
	Link to Parent Order			OBR.29		Parent Test Code								20		O
*/
hl7.OBX := []
hl7.OBX.map := {2:"Obs type",3:"resCode^resName",5:"resValue",6:"resUnits"}
/*	This is a required segment. It contains the values corresponding to the results.
	Sequence Number					OBX.01		Serial Number starting from 1					4		R
	Type Value						OBX.02		CE�Coded entry NM�Num ST�String TX-Text			2		R
	Observation Identifier			OBX.03		Result code 20^Result Name 255					275		R
	Observation Sub-ID				OBX.04		Not Supported											N
	Observation Value				OBX.05		Observation Value								200		R
	Units							OBX.06		Units											15		R
	References Range				OBX.07		Reference Range									40		R
	Abnormal Flags					OBX.08		L�Low H�High LL�Alert low HH�Alert high			2		R
	Probability						OBX.09		Not Supported											N
	Nature of Abnormal Test			OBX.10		Not Supported											N
	Observation Result Status		OBX.11		P�Prelim F�Final								1		O
	Date Last Normal				OBX.12		Not Supported											N
	User Defined Access Checks		OBX.13		Not Supported											N
	Date/Time of the Observation	OBX.14														26		O
	Producer�s ID					OBX.15		CODE											60		O
*/
hl7.NTE := []
hl7.NTE.map := {3:"Comment"}
/*	The Notes and Comments (NTE) segment contains notes and comments for the lab results and it is an optional segment.
	Segment Type ID					NTE.00		NTE												3		R
	Sequence Number					NTE.01		Serial Number starting from 1					4		R
	Source and Comment				NTE.02		Not Supported											N
	Comment							NTE.03		Comment											255		R
*/
	
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
