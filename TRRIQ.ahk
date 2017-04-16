/*	TRRIQ - The Rhythm Recording Interpretation Query
	Converts file
		Drag-and-drop onto window
		Monitor folder for changes
	Inputs a text file
		Probably converted from PDF using XPDF's "PDFtoTEXT"
		Use the -layout or -table option
		Only need the first 1-2 pages
	Identifies type of report:
		ZioPatch Holter
		LifeWatch (or other) Holter
	Extracts salient data
	Generates report using mail merge or template in Word
	Sends report to HIM
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
FileInstall, pdftotext.exe, pdftotext.exe
FileInstall, pdftk.exe, pdftk.exe
FileInstall, libiconv2.dll, libiconv2.dll

SplitPath, A_ScriptDir,,fileDir
user := A_UserName
IfInString, fileDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	chip := httpComm("full")
	FileDelete, .\Chipotle\currlist.xml
	FileAppend, % chip, .\Chipotle\currlist.xml
	isAdmin := true
	holterDir := ".\Holter PDFs\"
	importFld := ".\Import\"
	chipDir := ".\Chipotle\"
	eventlog(">>>>> Started in DEVT mode.")
} else {
	isAdmin := false
	holterDir := "..\Holter PDFs\"
	importFld := "..\Import\"
	chipDir := "\\childrens\files\HCChipotle\"
	eventlog(">>>>> Started in PROD mode.")
}

/*	Read outdocs.csv for Cardiologist and Fellow names 
*/
Docs := Object()
tmpChk := false
Loop, Read, %chipDir%outdocs.csv
{
	tmp := tmp0 := tmp1 := tmp2 := tmp3 := tmp4 := ""
	tmpline := A_LoopReadLine
	StringSplit, tmp, tmpline, `, , `"
	if ((tmp1="Name") or (tmp1="end") or !(tmp1)) {					; header, end, or blank lines
		continue
	}
	if (tmp4="group") {												; skip group names
		continue
	}
	if (tmp2="" and tmp3="" and tmp4="") {							; Fields 2,3,4 blank = new group
		tmpGrp := tmp1
		tmpIdx := 0
		tmpIdxG += 1
		outGrps.Insert(tmpGrp)
		continue
	}
	if !(tmp4~="i)(seattlechildrens.org)|(washington.edu)") {		; skip non-SCH or non-UW providers
		continue
	}
	tmpIdx += 1
	StringSplit, tmpPrv, tmp1, %A_Space%`"
	;tmpPrv := substr(tmpPrv1,1,1) . ". " . tmpPrv2					; F. Last
	tmpPrv := tmpPrv2 ", " tmpPrv1									; Last, First
	Docs[tmpGrp,tmpIdx]:=tmpPrv
	Docs[tmpGrp ".eml",tmpIdx] := tmp4
}

y := new XML(chipDir "currlist.xml")

;~ siteVals := {"CRD":"Seattle","EKG":"EKG lab","ECO":"ECHO lab","CRDBCSC":"Bellevue","CRDEVT":"Everett","CRDTAC":"Tacoma","CRDTRI":"Tri Cities","CRDWEN":"Wenatchee","CRDYAK":"Yakima"}
demVals := ["MRN","Account Number","DOB","Sex","Loc","Provider"]						; valid field names for parseClip()

if !(phase) {
	phase := CMsgBox("Which task?",""
		, "*&Upload LifeWatch Holter|"
		. "&Process PDF"
		,"Q","")
	;~ phase := CMsgBox("Which task?","","*&Register Preventice|&Process PDF file(s)","Q","")
}
if (instr(phase,"LifeWatch")) {
	eventlog("Start LifeWatch upload process.")
	Loop 
	{
		ptDem := Object()
		gosub fetchGUI								; Draw input GUI
		gosub fetchDem								; Grab demographics from CIS until accept
		gosub zybitSet								; Fill in Zybit demographics
	}
	ExitApp
}
;~ if (instr(phase,"preventice")) {
	;~ Loop
	;~ {
		;~ ptDem := Object()
		;~ gosub fetchGUI
		;~ gosub fetchDem
		;~ gosub webFill
	;~ }
	;~ ExitApp
;~ }
if (instr(phase,"PDF")) {
	eventlog("Start PDF folder scan.")
	holterLoops := 0								; Reset counters
	holtersDone := 
	loop, %holterDir%*.pdf							; Process all PDF files in holterDir
	{
		fileNam := RegExReplace(A_LoopFileName,"i)\.pdf")				; fileNam is name only without extension, no path
		fileIn := A_LoopFileFullPath									; fileIn has complete path \\childrens\files\HCCardiologyFiles\EP\HoltER Database\Holter PDFs\steve.pdf
		FileGetTime, fileDt, %fileIn%, C								; fildDt is creatdate/time 
		if (substr(fileDt,-5,2)<4) {									; skip files with creation TIME 0000-0359 (already processed)
			eventlog("Skipping file """ fileNam ".pdf"", already processed.")	; should be more resistant to DST. +0100 or -0100 will still be < 4
			continue
		}
		eventlog("Processing """ fileNam ".pdf"".")
		gosub MainLoop													; process the PDF
		if (fetchQuit=true) {											; [x] out of fetchDem means skip this file
			continue
		}
		if !IsObject(ptDem) {											; bad file, never acquires demographics
			continue
		}
		FileDelete, %fileIn%
		holterLoops++													; increment counter for processed counter
		holtersDone .= A_LoopFileName "->" filenameOut ".pdf`n"			; add to report
	}
	MsgBox % "Holters processed (" holterLoops ")`n" holtersDone
	/* Consider asking if complete. The MA's appear to run one PDF at a time, despite the efficiency loss.
	*/
}
eventlog("<<<<< Session end.")
ExitApp

FetchDem:
{
	mdX := Object()										; clear Mouse Demographics X,Y coordinate arrays
	mdY := Object()
	getDem := true
	while (getDem) {									; Repeat until we get tired of this
		clipboard :=
		ClipWait, 2
		if !ErrorLevel {								; clipboard has data
			clk := parseClip(clipboard)
			if !ErrorLevel {															; parseClip {field:value} matches valid data
				MouseGetPos, mouseXpos, mouseYpos, mouseWinID, mouseWinClass, 2			; put mouse coords into mouseXpos and mouseYpos, and associated winID
				if (clk.field = "Provider") {
					if (clk.value~="[[:alpha:]]+.*,.*[[:alpha:]]+") {						; extract provider.value to LAST,FIRST (strip MD, PHD, MI, etc)
						tmpPrv := strX(clk.value,,1,0, ",",1,1) ", " strX(clk.value,",",1,2, " ",1,1)
						eventlog("MouseGrab provider " tmpPrv ".")
					} else {
						tmpPrv :=
						eventlog("MouseGrab provider empty.")
					}
					if ((ptDem.Provider) && (tmpPrv)) {												; Provider already exists
						MsgBox, 4148
							, Provider already exists
							, % "Replace " ptDem.Provider "`n with `n" tmpPrv "?"
						IfMsgBox, Yes													; Check before replacing
						{
							eventlog("Replacing provider """ ptDem.Provider """ with """ tmpPrv """.")
							ptDem.Provider := tmpPrv
						}
					} else if (tmpPrv) {												; Otherwise populate ptDem.Provider if tmpPrv exists
						ptDem.Provider := tmpPrv										; but leave ptDem.Provider alone if tmpPrv null
						eventlog("MouseGrab provider empty --> " tmpPrv ".")
					}
					mdX[4] := mouseXpos													; demographics grid[4,1]
					mdY[1] := mouseYpos
					mdProv := true														; we have got Provider
					WinGetTitle, mdTitle, ahk_id %mouseWinID%
					gosub getDemName													; extract patient name, MRN from window title 
					
				}																		;(this is why it must be sister or parent VM).
				if (clk.field = "Account Number") {
					ptDem["Account Number"] := clk.value
					eventlog("MouseGrab Account Number " clk.value ".")
					mdX[1] := mouseXpos													; demographics grid[1,3]
					mdY[3] := mouseYpos
					mdAcct := true														; we have got Acct Number
					WinGetTitle, mdTitle, ahk_id %mouseWinID%
					gosub getDemName													; extract patient name, MRN
				}
				if (mdProv and mdAcct) {												; we have both critical coordinates
					mdXd := (mdX[4]-mdX[1])/3											; determine delta X between columns
					mdX[2] := mdX[1]+mdXd												; determine remaining column positions
					mdX[3] := mdX[2]+mdXd
					mdY[2] := mdY[1]+(mdY[3]-mdY[1])/2									; determine remaning row coordinate
					/*	possible to just divide the window width into 6 columns
						rather than dividing the space into delta X ?
					*/
					Gui, fetch:hide
					ptDem["MRN"] := mouseGrab(mdX[1],mdY[2]).value						; grab remaining demographic values
					ptDem["DOB"] := mouseGrab(mdX[2],mdY[2]).value
					ptDem["Sex"] := substr(mouseGrab(mdX[3],mdY[1]).value,1,1)
					eventlog("MouseGrab other fields. MRN=" ptDem["MRN"] " DOB=" ptDem["DOB"] " Sex=" ptDem["Sex"] ".")
					
					tmp := mouseGrab(mdX[3],mdY[3])										; grab Encounter Type field
					ptDem["Type"] := tmp.value
					if (ptDem["Type"]="Outpatient") {
						ptDem["Loc"] := mouseGrab(mdX[3]+mdXd*0.5,mdY[2]).value			; most outpatient locations are short strings, click the right half of cell to grab location name
					} else {
						ptDem["Loc"] := tmp.loc
					}
					if !(ptDem["EncDate"]) {											; EncDate will be empty if new upload or null in PDF
						ptDem["EncDate"] := tmp.date
						ptDem["Hookup time"] := tmp.time
					}
					mdProv := false														; processed demographic fields,
					mdAcct := false														; so reset check bits
					Gui, fetch:show
					eventlog("MouseGrab other fields."
						. " Type=" ptDem["Type"] " Loc=" ptDem["Loc"]
						. " EncDate=" ptDem["EncDate"] " EncTime=" ptDem["Hookup time"] ".")
				}
			}
			gosub fetchGUI							; Update GUI with new info
		}
	}
	return
}

mouseGrab(x,y) {
/*	Double click mouse coordinates x,y to grab cell contents
	Process through parseClip to validate
	Return the value portion of parseClip
*/
	BlockInput, On																		; Prevent extraneous input
	MouseMove, %x%, %y%, 0																; Goto coordinates
	Click 2																				; Double-click
	BlockInput, Off																		; Permit input again
	sleep 100
	ClipWait																			; sometimes there is delay for clipboard to populate
	clk := parseClip(clipboard)															; get available values out of clipboard
	return clk																			; Redundant? since this is what parseClip() returns
}

parseClip(clip) {
/*	If clip matches "val1:val2" format, and val1 in demVals[], return field:val
	If clip contains proper Encounter Type ("Outpatient", "Inpatient", "Observation", etc), return Type, Date, Time
*/
	global demVals
	
	StringSplit, val, clip, :															; break field into val1:val2
	if (ObjHasValue(demVals, val1)) {													; field name in demVals, e.g. "MRN","Account Number","DOB","Sex","Loc","Provider"
		return {"field":val1
				, "value":val2}
	}
	
	dt := strX(clip," [",1,2, "]",1,1)													; get date
	dd := parseDate(dt).YYYY . parseDate(dt).MM . parseDate(dt).DD
	if (clip~="Outpatient\s\[") {														; Outpatient type
		return {"field":"Type"
				, "value":"Outpatient"
				, "loc":"Outpatient"
				, "date":dt
				, "time":parseDate(dt).time}
	}
	if (clip~="Inpatient|Observation\s\[") {											; Inpatient types
		return {"field":"Type"
				, "value":"Inpatient"
				, "loc":"Inpatient"
				, "date":""}															; can span many days, return blank
	}
	if (clip~="Day Surg.*\s\[") {														; Day Surg type
		return {"field":"Type"
				, "value":"Day Surg"
				, "loc":"SurgCntr"
				, "date":dt}
	}
	if (clip~="Emergency") {															; Emergency type
		return {"field":"Type"
				, "value":"Emergency"
				, "loc":"Emergency"
				, "date":dt}
	}
	return Error																		; Anything else returns Error
}

getDemName:
{
	if (RegExMatch(mdTitle, "i)\s\-\s\d{6,7}\s(Opened by)")) {							; Match window title "LAST, FIRST - 12345678 Opened by Chun, Terrence U, MD"
		ptDem["nameL"] := strX(mdTitle,,1,0, ",",1,1)									; and parse the name
		ptDem["nameF"] := strX(mdTitle,",",1,2, " ",1,1)
	}
	return
}

fetchGUI:
{
	fYd := 30,	fXd := 90									; fetchGUI delta Y, X
	fX1 := 12,	fX2 := fX1+fXd								; x pos for title and input fields
	fW1 := 80,	fW2 := 190									; width for title and input fields
	fH := 20												; line heights
	fY := 10												; y pos to start
	EncNum := ptDem["Account Number"]						; we need these non-array variables for the Gui statements
	encDT := parseDate(ptDem.EncDate).YYYY . parseDate(ptDem.EncDate).MM . parseDate(ptDem.EncDate).DD
	demBits := 0											; clear the error check
	fTxt := "	To auto-grab demographic info:`n"
		.	"		1) Double-click Account Number #`n"
		.	"		2) Double-click Provider"
	Gui, fetch:Destroy
	Gui, fetch:+AlwaysOnTop
	Gui, fetch:Add, Text, % "x" fX1 , % fTxt	
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd*2) " w" fW1 " h" fH " c" fetchValid("nameF","[a-z]"), First
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " cDefault", % ptDem["nameF"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("nameL","[a-z]"), Last
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["nameL"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("MRN","\d{6,7}",1), MRN
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " cDefault", % ptDem["MRN"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("DOB","\d{1,2}/\d{1,2}/\d{2,4}",1), DOB
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " cDefault", % ptDem["DOB"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Date placed
	Gui, fetch:Add, DateTime, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " vEncDt CHOOSE" encDT, MM/dd/yyyy
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("Loc","i)[a-z]+",1), Location
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " cDefault", % ptDem["Loc"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("Type","i)[a-z]+",1), Type
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " cDefault", % ptDem["Type"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("Account Number","\d{8}",1), Encounter #
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " vEncNum" " cDefault", % encNum
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" ((!(checkCrd(ptDem.Provider).fuzz=0)||!(ptDem.Provider))?"Red":"Default"), Ordering MD
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH  " cDefault", % ptDem["Provider"]
	Gui, fetch:Add, Button, % "x" fX1+10 " y" (fY += fYD) " h" fH+10 " w" fW1+fW2 " gfetchSubmit " ((demBits)?"Disabled":""), Submit!
	Gui, fetch:Show, AutoSize, Enter Demographics
	return
}

fetchValid(field,rx,neg:=0) {
/* 	checks regex(rx) for ptDem[field] 
 *	if neg, gives opposite result
 *	any negative result returns demBit
 */
	global ptDem, demBits
	if !(ptDem[field]) {
		demBits := 1
		return "Red"
	}
	res := (ptDem[field]~=rx)
	if (neg) {
		demBits := !(res)
		return ((res)?"Default":"Red")
	} else {
		demBits := (res)
		return ((res)?"Red":"Default")
	}
}

fetchGuiClose:
	Gui, fetch:destroy
	getDem := false																	; break out of fetchDem loop
	fetchQuit := true
	eventlog("Manual [x] out of fetchDem.")
Return

fetchSubmit:
{
/* some error checking
	Check for required elements
	Error check and normalize Ordering MD name
	Check for Lifewatch exe
	Fill Lifewatch data and submit
	The repeat the cycle
demVals := ["MRN","Account Number","DOB","Sex","Loc","Provider"]
*/
	Gui, fetch:Submit
	Gui, fetch:Destroy
	
	if (instr(ptDem.Provider," ") && !instr(ptDem.Provider,",")) {				; somehow string passed in wrong order
		tmp := trim(ptDem.Provider)
		tmpF := strX(tmp,"",1,0, " ",1,1)
		tmpL := strX(tmp," ",1,1, "",1,0)
		ptDem.Provider := tmpL ", " tmpF
	}
	matchProv := checkCrd(ptDem.Provider)
	if !(ptDem.Provider) {														; no provider? ask!
		gosub getMD
		eventlog("New provider field " ptDem.Provider ".")
	} else if (matchProv.fuzz > 0.10) {							; Provider not recognized
		eventlog(ptDem.Provider " not recognized (" matchProv.fuzz ").")
		if (ptDem.Type~="i)(Inpatient|Observation|Emergency|Day Surg)") {
			gosub assignMD														; Inpt, ER, DaySurg, we must find who recommended it from the Chipotle schedule
			eventlog(ptDem.Type " location. Provider assigned to " ptDem.Provider ".")
		} else {
			gosub getMD															; Otherwise, ask for it.
			eventlog("Provider set to " ptDem.Provider ".")
		}
	} else {													; Provider recognized
		eventlog(ptDem.Provider " matches " matchProv.Best " (" matchProv.fuzz ").")
		ptDem.Provider := matchProv.Best
	}
	ptDem["Account Number"] := EncNum											; make sure array has submitted EncNum value
	FormatTime, EncDt, %EncDt%, MM/dd/yyyy										; and the properly formatted date 06/15/2016
	ptDem.EncDate := EncDt
	ptDemChk := (ptDem["nameF"]~="i)[A-Z\-]+") && (ptDem["nameL"]~="i)[A-Z\-]+") 					; valid names
			&& (ptDem["mrn"]~="\d{6,7}") && (ptDem["Account Number"]~="\d{8}") 						; valid MRN and Acct numbers
			&& (ptDem["DOB"]~="[0-9]{1,2}/[0-9]{1,2}/[1-2][0-9]{3}") && (ptDem["Sex"]~="[MF]") 		; valid DOB and Sex
			&& (ptDem["Loc"]) && (ptDem["Type"])													; Loc and type is not null
			&& (ptDem["Provider"]~="i)[a-z]+") && (ptDem["EncDate"])								; prov any string, encDate not null
	if !(ptDemChk) {																	; all data elements must be present, otherwise retry
		eventlog("Data incomplete."
			. ((ptDem["nameF"]) ? "" : " nameF")
			. ((ptDem["nameL"]) ? "" : " nameL")
			. ((ptDem["mrn"]) ? "" : " MRN")
			. ((ptDem["Account number"]) ? "" : " EncNum")
			. ((ptDem["DOB"]) ? "" : " DOB")
			. ((ptDem["Sex"]) ? "" : " Sex")
			. ((ptDem["Loc"]) ? "" : " Loc")
			. ((ptDem["Type"]) ? "" : " Type")
			. ((ptDem["EncDate"]) ? "" : " EncDate")
			. ((ptDem["Provider"]) ? "" : " Provider")
			. ".")
		MsgBox,, % "Data incomplete. Try again", % ""
			. ((ptDem["nameF"]) ? "" : "First name`n")
			. ((ptDem["nameL"]) ? "" : "Last name`n")
			. ((ptDem["mrn"]) ? "" : "MRN`n")
			. ((ptDem["Account number"]) ? "" : "Account number`n")
			. ((ptDem["DOB"]) ? "" : "DOB`n")
			. ((ptDem["Sex"]) ? "" : "Sex`n")
			. ((ptDem["Loc"]) ? "" : "Location`n")
			. ((ptDem["Type"]) ? "" : "Visit type`n")
			. ((ptDem["EncDate"]) ? "" : "Date Holter placed`n")
			. ((ptDem["Provider"]) ? "" : "Provider`n")
			. "`nREQUIRED!"
		;~ MsgBox % ""
			;~ . "First name " ptDem["nameF"] "`n"
			;~ . "Last name " ptDem["nameL"] "`n"
			;~ . "MRN " ptDem["mrn"] "`n"
			;~ . "Account number " ptDem["Account number"] "`n"
			;~ . "DOB " ptDem["DOB"] "`n"
			;~ . "Sex " ptDem["Sex"] "`n"
			;~ . "Location " ptDem["Loc"] "`n"
			;~ . "Visit type " ptDem["Type"] "`n"
			;~ . "Date Holter placed " ptDem["EncDate"] "`n"
			;~ . "Provider " ptDem["Provider"] "`n"
		gosub fetchGUI
		return
	}
	getDem := false																; done getting demographics
	Loop
	{
		if (ptDem.Indication) {													; loop until we have filled indChoices
			break
		}
		gosub indGUI
		WinWaitClose, Enter indications
	}
	eventlog("Indications entered.")
	return
}

indGUI:
{
	indOpts := ""
		. "Abnormal Electrocardiogram/Rhythm Strip" "|"
		. "Bradycardia" "|"
		. "Chest Pain" "|"
		. "Cyanosis" "|"
		. "Dizziness" "|"
		. "Electrolyte Disorder" "|"
		. "Failure to thrive" "|"
		. "Fever" "|"
		. "History of Cardiovascular Disease" "|"
		. "Hypertension" "|"
		. "Kawasaki Disease" "|"
		. "Medication requiring ECG surveillance" "|"
		. "Palpitations" "|"
		. "Premature Atrial Contractions (PAC's)" "|"
		. "Premature Ventricular Contractions (PVC's)" "|"
		. "Respiratory Distress" "|"
		. "Shortness of Breath" "|"
		. "Supraventricular Tachycardia (SVT)" "|"
		. "Syncope" "|"
		. "Tachycardia" "|"
		. "OTHER"
	Gui, ind:Destroy
	Gui, ind:+AlwaysOnTop
	Gui, ind:font, s12
	Gui, ind:Add, Text, , % "Enter indications: " ptDem["Indication"]
	Gui, ind:Add, ListBox, r12 vIndChoices 8, %indOpts%
	Gui, ind:Add, Button, gindSubmit, Submit
	Gui, ind:Show, Autosize, Enter indications
	return
}

indClose:
ExitApp

indSubmit:
{
	Gui, ind:Submit
	if InStr(indChoices,"OTHER",Yes) {
		InputBox, indOther, Other, Enter other indication
		indChoices := RegExReplace(indChoices,"OTHER", "OTHER - " indOther)
	}
	ptDem["Indication"] := indChoices
	return
}

webFill:
{
	MsgBox Fill a form
	
	return
}

zybitSet:
{
	Loop
	{
		if (zyWinId := WinExist("ahk_exe ZybitRemote.exe")) {
			break
		}
		MsgBox, 262193, Inject demographics, Must run Zybit Holter program!
		IfMsgBox Cancel
			ExitApp
	}
	Loop
	{
		if (zyNewId := WinExist("New Patient - Demographics")) {
			break
		}
		MsgBox, 262192, Start NEW patient, Click OK when ready to inject demographic information
	}
	zyVals := {"Edit1":ptDem["nameL"],"Edit2":ptDem["nameF"]
				,"Edit4":ptDem["Sex"],"Edit5":ptDem["DOB"]
				,"Edit6":ptDem["mrn"],"Edit8":ptDem["Account number"]
				,"Edit7":ptDem["loc"]
				,"Edit9":indChoices
				,"Edit11":ptDem["Provider"],"Edit12":user }
	
	zybitFill(zyNewId,zyVals)
	
	; Log the entry?
	
	return
}

zybitFill(win,fields) {
	WinActivate, ahk_id %win%
	for key,val in fields
	{
		ControlSetText, %key%, %val%, ahk_id %win%
	}
	return
}

MainLoop:
{
/*	This main loop accepts a %fileIn% filename,
 *	determines the filetype based on header contents,
 *	concatenates the CSV strings of header (fileOut1) and values (fileOut2)
 *	into a single file (fileOut),
 *	move around the temp, CSV, and PDF files.
 */
	RunWait, pdftotext.exe -l 2 -table -fixed 3 "%fileIn%" temp.txt					; convert PDF pages 1-2 to txt file
	newTxt:=""																		; clear the full txt variable
	FileRead, maintxt, temp.txt														; load into maintxt
	StringReplace, newtxt, maintxt, `r`n`r`n, `r`n, All
	FileDelete tempfile.txt															; remove any leftover tempfile
	FileAppend %newtxt%, tempfile.txt												; create new tempfile with newtxt result
	FileMove tempfile.txt, .\tempfiles\%fileNam%.txt								; move a copy into tempfiles for troubleshooting
	eventlog("tempfile.txt -> " fileNam ".txt")

	blocks := Object()																; clear all objects
	fields := Object()
	fldval := {}
	labels := Object()
	blk := Object()
	blk2 := Object()
	ptDem := Object()
	chk := Object()
	matchProv := Object()
	fileOut := fileOut1 := fileOut2 := ""
	summBl := summ := ""
	
	if ((newtxt~="i)Philips|Lifewatch") && instr(newtxt,"Holter")) {					; Processing loop based on identifying string in newtxt
		gosub Holter_LW
	} else if (instr(newtxt,"Preventice") && instr(newtxt,"H3Plus")) {
		gosub Holter_Pr
	} else if (RegExMatch(newtxt,"i)zio.*xt.*patch")) {
		gosub Zio
	} else if (InStr(newtxt,"TRANSTELEPHONIC ARRHYTHMIA")) {
		gosub Event_LW
	} else if (instr(newtxt,"Preventice") && instr(newtxt,"End of Service Report")) {
		gosub Event_BGH
	} else {
		eventlog(fileNam " bad file.")
		MsgBox No match!
		return
	}
	if (fetchQuit=true) {																	; exited demographics fetchGUI
		return																				; so skip processing this file
	}
	gosub epRead																			; find out which EP is reading today
	/*	Output the results and move files around
	*/
	fileOut1 .= (substr(fileOut1,0,1)="`n") ?: "`n"											; make sure that there is only one `n 
	fileOut2 .= (substr(fileOut2,0,1)="`n") ?: "`n"											; on the header and data lines
	fileout := fileOut1 . fileout2															; concatenate the header and data lines
	tmpDate := parseDate(fldval["Test_Date"])												; get the study date
	filenameOut := fldval["MRN"] " " fldval["Name_L"] " " tmpDate.MM "-" tmpDate.DD "-" tmpDate.YYYY
	tmpFlag := tmpDate.YYYY . tmpDate.MM . tmpDate.DD . "020000"
	FileDelete, .\tempfiles\%fileNameOut%.csv												; clear any previous CSV
	FileAppend, %fileOut%, .\tempfiles\%fileNameOut%.csv									; create a new CSV
	FileCopy, .\tempfiles\%fileNameOut%.csv, %importFld%*.*, 1								; create a copy of CSV in tempfiles
	FileCopy, %fileIn%, %holterDir%Archive\%filenameOut%.pdf, 1								; move the PDF to holterDir
	FileMove, %fileIn%sh.pdf, %holterDir%%filenameOut%-short.pdf, 1							; move the shortened PDF, if it exists
	FileSetTime, tmpFlag, %holterDir%Archive\%filenameOut%.pdf, C							; set the time of PDF in holterDir to 020000 (processed)
	FileSetTime, tmpFlag, %holterDir%%filenameOut%-short.pdf, C
	eventlog("Move files " filenameOut)
Return
}

getMD:
{
	Gui, fetch:Hide
	InputBox, ed_Crd, % "Enter responsible cardiologist"						; no call schedule for that day, must choose
	Gui, fetch:Show
	if (ed_Crd="")
		return
	tmpCrd := checkCrd(ed_Crd)
	if (tmpCrd.fuzz=0) {										; Perfect match found
		ptDem.Provider := tmpCrd.best
	} else {													; less than perfect
		MsgBox, 262180, Cardiologist
			, % "Did you mean: " tmpCrd.best "?`n`n`n"
		IfMsgBox, Yes
		{
			ptDem.Provider := tmpCrd.best
		}
	}
	eventlog("Cardiologist " ptDem.Provider " entered.")
	return
}	

assignMD:
{
	if !(ptDem.EncDate) {														; must have a date to figure it out
		return
	}
	encDT := parseDate(ptDem.EncDate).YYYY . parseDate(ptDem.EncDate).MM . parseDate(ptDem.EncDate).DD
	inptMD := checkCrd(ptDem.Provider) 
	if (inptMD.fuzz=0) {														; attg is Crd
		ptDem.Loc := "Inpatient"												; set Loc so it won't fail
		return
	} 
	if (ymatch := y.selectSingleNode("//call[@date='" encDT "']/Ward_A").text) {
		inptMD := checkCrd(strX(ymatch," ",1,1) ", " strX(ymatch,"",1,0," ",1,1))
		if (inptMD.fuzz=0) {													; on-call Cards that day 
			ptDem.Loc := "Inpatient"
			ptDem.Provider := inptMD.best
			eventlog("Cardiologist autoselected " ptDem.Provider )
			return
		}
	}
	gosub getMD																	; when all else fails, ask
	ptDem.Loc := "Inpatient"
return
}

epRead:
{
	FileGetTime, dlDate, %fileIn%
	FormatTime, dlDay, %dlDate%, dddd
	if (dlDay="Friday") {
		dlDate += 3, Days
	}
	FormatTime, dlDate, %dlDate%, yyyyMMdd
	
	RegExMatch(y.selectSingleNode("//call[@date='" dlDate "']/EP").text, "Oi)(Chun)|(Salerno)|(Seslar)", ymatch)
	if !(ymatch := ymatch.value()) {
		ymatch := epMon ? epMon : cmsgbox("Electronic Forecast not complete","Which EP on Monday?","Chun|Salerno|Seslar","Q")
		epMon := ymatch
		eventlog("Reading EP assigned to " epMon ".")
	}
	
	if (RegExMatch(fldval["ordering"], "Oi)(Chun)|(Salerno)|(Seslar)", epOrder))  {
		ymatch := epOrder.value()
	}
	
	FormatTime, ma_date, A_Now, MM/dd/yyyy
	fileOut1 .= ",""EP_read"",""EP_date"",""MA"",""MA_date"""
	fileOut2 .= ",""" ymatch """,""" niceDate(dlDate) """,""" user """,""" ma_date """"
return
}

Holter_LW:
{
	eventlog("Holter_LW")
	monType := "H"
	
	demog := columns(newtxt,"PATIENT DEMOGRAPHICS","Heart Rate Data",,"Reading Physician")
	holtVals := columns(newtxt,"Medications","INTERPRETATION",,"Total VE Beats")
	
	gosub checkProcLW											; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	/* Holter PDF is valid. OK to process.
	 * Pulls text between field[n] and field[n+1], place in labels[n] name, with prefix "dem-" etc.
	 */
	fields[1] := ["Last Name", "First Name", "Middle Initial", "ID Number", "Date Of Birth", "Sex"
		, "Source", "Billing Code", "Recorder Format", "Pt\.? Home\s*(Phone)?\s*#?", "Hookup Tech", "Pacemaker\s*Y/N.", "Medications"
		, "Physician", "Scanned By", "Reading Physician"
		, "Test Date", "Analysis Date", "Hookup Time", "Recording Time", "Analysis Time", "Reason for Test", "(Group|user field)"]
	labels[1] := ["Name_L", "Name_F", "Name_M", "MRN", "DOB", "Sex"
		, "Site", "Billing", "Device_SN", "VOID1", "Hookup_tech", "VOID2", "Meds"
		, "Ordering", "Scanned_by", "Reading"
		, "Test_date", "Scan_date", "Hookup_time", "Recording_time", "Analysis_time", "Indication", "VOID3"]
	fieldvals(demog,1,"dem")
	
	fields[2] := ["Total Beats", "Min HR", "Avg HR", "Max HR", "Heart Rate Variability"]
	labels[2] := ["Total_beats", "Min", "Avg", "Max", "HRV"]
	fieldvals(strX(holtVals,"Heart Rate Data",1,0,"Heart Rate Variability",1,0,nn),2,"hrd")
	
	fields[3] := ["Total VE Beats", "Vent Runs", "Beats", "Longest", "Fastest", "Triplets", "Couplets", "Single/Interp PVC", "R on T", "Single/Late VE's", "Bi/Trigeminy", "Supraventricular Ectopy"]
	labels[3] := ["Total", "Runs", "Beats", "Longest", "Fastest", "Triplets", "Couplets", "SinglePVC_InterpPVC", "R_on_T", "SingleVE_LateVE", "Bigem_Trigem", "SVE"]
	fieldvals(strX(holtVals,"Ventricular Ectopy",nn,0,"Supraventricular Ectopy",1,0,nn),3,"ve")

	fields[4] := ["Total SVE Beats", "Atrial Runs", "Beats", "Longest", "Fastest", "Atrial Pairs", "Drop/Late", "Longest R-R", "Single PAC's", "Bi/Trigeminy", "Atrial Fibrillation"]
	labels[4] := ["Total", "Runs", "Beats", "Longest", "Fastest", "Pairs", "Drop_Late", "LongRR", "Single", "Bigem_Trigem", "AF"]
	fieldvals(strX(holtVals,"Supraventricular Ectopy",nn-23,0,"Atrial Fibrillation",1,0,nn),4,"sve")
	
	tmp := strX(RegExReplace(newtxt,"i)technician.*comments?:","TECH COMMENT:"),"TECH COMMENT:",1,13,"",1,0)
	StringReplace, tmp, tmp, .`n , .%A_Space% , All
	fileout1 .= """INTERP"""
	fileout2 .= """" cleanspace(trim(tmp," `n")) """"
	fileOut1 .= ",""Mon_type"""
	fileOut2 .= ",""Holter"""
	
	ShortenPDF("FULL DISCLOSURE")

return
}

Holter_Pr:
{
	eventlog("Holter_Pr")
	monType := "PR"
	
	demog := columns(newtxt,"Patient Information","Scan Criteria",1,"Date Processed")
	sumStat := columns(newtxt,"Summary Statistics","Rate Statistics",1,"Recording Duration","Analyzed Data")
	rateStat := columns(newtxt,"Rate Statistics","Supraventricular Ectopy",,"Tachycardia/Bradycardia") "#####"
	ectoStat := columns(newtxt,"Supraventricular Ectopy","ST Deviation",,"Ventricular Ectopy")
	pauseStat := columns(newtxt,"Pauses","Comment",,"\# RRs")
	
	gosub checkProcPR											; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	/* Holter PDF is valid. OK to process.
	 * Pulls text between field[n] and field[n+1], place in labels[n] name, with prefix "dem-" etc.
	 */
	fields[1] := ["Name", "ID #", "Second ID", "Date Of Birth", "Age", "Sex"
		, "Referring Physician", "Indications", "Medications", "Analyst", "Hookup Tech"
		, "Date Recorded", "Date Processed", "Scan Number", "Recorder", "Recorder No", "Hookup time", "Location", "Acct num"]
	labels[1] := ["Name", "MRN", "VOID_ID", "DOB", "VOID_Age", "Sex"
		, "Ordering", "Indication", "Meds", "Scanned_by", "Hookup_tech"
		, "Test_date", "Scan_date", "Scan_num", "Recorder", "Device_SN", "Hookup_time", "Site", "Billing"]
	fieldvals(demog,1,"dem")
	
	fields[2] := ["Total QRS", "Recording Duration", "Analyzed Data"]
	labels[2] := ["Total_beats", "dem:Recording_time", "dem:Analysis_time"]
	fieldvals(sumStat,2,"hrd")
	
	fields[3] := ["Min Rate", "Max Rate", "Mean Rate", "Tachycardia/Bradycardia"
		, "Longest Tachycardia", "Fastest Tachycardia", "Longest Bradycardia", "Slowest Bradycardia","#####"]
	labels[3] := ["Min", "Max", "Avg", "VOID_tb"
		, "Longest_tachy", "Fastest", "Longest_brady", "Slowest","VOID_br"]
	fieldvals(rateStat,3,"hrd")
	
	SveStat := strVal(ectoStat,"Supraventricular Ectopy","Ventricular Ectopy")
	fields[4] := ["Singles", "Couplets", "Runs", "Total","RR Variability"]
	labels[4] := ["Beats", "Pairs", "Runs", "Total","RR_var"]
	if (SveStat~="i)Fastest.*Longest") {
		fields[4].Insert(4,"Fastest Run","Longest Run")
		labels[4].Insert(4,"Fastest","Longest")
	}
	fieldvals(SveStat,4,"sve")
	
	VeStat := strVal(ectoStat "#####","\bVentricular Ectopy","#####")
	fields[5] := ["Singles", "Couplets", "Runs", "R on T", "Total"]
	labels[5] := ["Beats", "Couplets", "Runs", "R on T", "Total"]
	if (VeStat~="i)Fastest.*Longest") {
		fields[5].Insert(4,"Fastest Run","Longest Run")
		labels[5].Insert(4,"Fastest","Longest")
	}
	fieldvals(VeStat,5,"ve")

	fields[6] := ["Longest RR","\# RR.*\..* sec"]
	labels[6] := ["LongRR","Pauses"]
	fieldvals(pauseStat,6,"sve")
	
	LWify()
	tmp := strVal(newtxt,"COMMENT:","REVIEWING PHYSICIAN")
	StringReplace, tmp, tmp, .`n , .%A_Space% , All
	fileout1 .= """INTERP"""
	fileout2 .= """" cleanspace(trim(tmp," `n")) """"
	fileOut1 .= ",""Mon_type"""
	fileOut2 .= ",""Holter"""
	
	ShortenPDF("i)60\s+sec/line")

return
}

LWify() {
	global fileout1, fileout2
	lwfields := {"dem-Name_L":"", "dem-Name_F":"", "dem-Name_M":"", "dem-MRN":"", "dem-DOB":"", "dem-Sex":"NA"
				, "dem-Site":"", "dem-Billing":"", "dem-Device_SN":"", "dem-VOID1":"", "dem-Hookup_tech":""
				, "dem-VOID2":"", "dem-Meds":"NA", "dem-Ordering":"", "dem-Ordering_grp":"", "dem-Ordering_eml":""
				, "dem-Scanned_by":"", "dem-Reading":"", "dem-Test_date":"", "dem-Scan_date":"", "dem-Hookup_time":""
				, "dem-Recording_time":"", "dem-Analysis_time":"", "dem-Indication":"", "dem-VOID3":""
				, "hrd-Total_beats":"0", "hrd-Min":"0", "hrd-Min_time":"", "hrd-Avg":"0", "hrd-Max":"0", "hrd-Max_time":"", "hrd-HRV":""
				, "ve-Total":"0", "ve-Total_per":"0", "ve-Runs":"0", "ve-Beats":"0", "ve-Longest":"0", "ve-Longest_time":""
				, "ve-Fastest":"0", "ve-Fastest_time":"", "ve-Triplets":"0", "ve-Couplets":"0", "ve-SinglePVC":"0", "ve-InterpPVC":"0"
				, "ve-R_on_T":"0", "ve-SingleVE":"0", "ve-LateVE":"0", "ve-Bigem":"0", "ve-Trigem":"0", "ve-SVE":"0"
				, "sve-Total":"0", "sve-Total_per":"0", "sve-Runs":"0", "sve-Beats":"0", "sve-Longest":"0", "sve-Longest_time":""
				, "sve-Fastest":"0", "sve-Fastest_time":"", "sve-Pairs":"0", "sve-Drop":"0", "sve-Late":"0"
				, "sve-LongRR":"0", "sve-LongRR_time":"", "sve-Single":"0", "sve-Bigem":"0", "sve-Trigem":"0", "sve-AF":"0"}
	
	lwtabs := "dem-Name_L	dem-Name_F	dem-Name_M	dem-MRN	dem-DOB	dem-Sex	dem-Site	dem-Billing	dem-Device_SN	dem-VOID1	"
		. "dem-Hookup_tech	dem-VOID2	dem-Meds	dem-Ordering	dem-Ordering_grp	dem-Ordering_eml	dem-Scanned_by	"
		. "dem-Reading	dem-Test_date	dem-Scan_date	dem-Hookup_time	dem-Recording_time	dem-Analysis_time	dem-Indication	"
		. "dem-VOID3	hrd-Total_beats	hrd-Min	hrd-Min_time	hrd-Avg	hrd-Max	hrd-Max_time	hrd-HRV	ve-Total	ve-Total_per	"
		. "ve-Runs	ve-Beats	ve-Longest	ve-Longest_time	ve-Fastest	ve-Fastest_time	ve-Triplets	ve-Couplets	ve-SinglePVC	"
		. "ve-InterpPVC	ve-R_on_T	ve-SingleVE	ve-LateVE	ve-Bigem	ve-Trigem	ve-SVE	sve-Total	sve-Total_per	sve-Runs	"
		. "sve-Beats	sve-Longest	sve-Longest_time	sve-Fastest	sve-Fastest_time	sve-Pairs	sve-Drop	sve-Late	"
		. "sve-LongRR	sve-LongRR_time	sve-Single	sve-Bigem	sve-Trigem	sve-AF"
	
	field := Object()
	
	; Parse label lines in CSV
	loop, Parse, fileOut1, CSV
	{
		field.push(A_LoopField)
	}
	; Parse values in CSV, match to equivalent labels
	loop, Parse, fileOut2, CSV
	{
		lwFields[field[A_index]] := A_LoopField
		if !objhaskey(lwFields,field[A_Index]) {
			MsgBox % field[A_Index]
		}
	}
	; Populate lwFields with equivalent named label values
	lwOut1 :=
	lwOut2 :=
	loop, parse, lwTabs, `t
	{
		fld := A_LoopField
		val := lwFields[fld]
		lwOut1 .= """" fld ""","
		lwOut2 .= """" val ""","
	}
	fileOut1 := lwOut1
	fileOut2 := lwOut2
	eventlog("LWify complete.")
return	
}

shortenPDF(find) {
	eventlog("ShortenPDF")
	global fileIn, winCons
	sleep 500
	ConsWin := WinExist("ahk_pid " winCons)								; get window ID

	loop, 100
	{
		FileGetSize, fullsize, tempfull.txt
		IfWinNotExist ahk_id %consWin% 
		{
			break
		}
		Progress, % A_index,% fullsize,Scanning full size PDF...%winCons%,%consWin%
		
		sleep 120
	}
	progress,100,fullsize, Shrinking PDF...
	FileRead, fulltxt, tempfull.txt
	filedelete, tempfull.txt
	findpos := RegExMatch(fulltxt,find)
	pgpos := instr(fulltxt,"Page ",,findpos-strlen(fulltxt))
	RegExMatch(fulltxt,"Oi)Page\s+(\d+)\s",pgs,pgpos)
	pgpos := pgs.value(1)
	RunWait, pdftk.exe "%fileIn%" cat 1-%pgpos% output "%fileIn%sh.pdf",,min
	progress, off
return	
}

CheckProcLW:
{
	eventlog("CheckProcLW")
	chk.Last := strVal(demog,"Last Name","First Name")						; NameL				must be [A-Z]
	chk.First := strVal(demog,"First Name","Middle Initial")				; NameF				must be [A-Z]
	chk.MRN := strVal(demog,"ID Number","Date of Birth")					; MRN
	chk.DOB := strVal(demog,"Date of Birth","Sex")							; DOB
	chk.Sex := strVal(demog,"Sex","Source")									; Sex
	chk.Loc := strVal(demog,"Source","Billing Code")						; Location			must be in SiteVals
	chk.Acct := strVal(demog,"Billing Code","Recorder Format")				; Billing code		must be valid number
	chk.Prov := strVal(demog,"Physician","Scanned By")						; Ordering MD
	chk.Date := strVal(demog,"Test Date","Analysis Date")					; Study date
	chk.Ind := strVal(demog,"Reason for Test","Group")						; Indication
	
	chkDT := parseDate(chk.Date)
	chkFilename := chk.MRN " * " chkDT.MM "-" chkDT.DD "-" chkDT.YYYY
	if FileExist(holterDir . "Archive\" . chkFilename . ".pdf") {
		FileDelete, %fileIn%
		eventlog(chk.MRN " PDF archive exists, deleting '" fileIn "'")
		fetchQuit := true
		return
	}
	
	Run , pdftotext.exe "%fileIn%" tempfull.txt,,min,wincons						; convert PDF all pages to txt file
	eventlog("Extracting full text.")
	
	Clipboard := chk.Last ", " chk.First														; fill clipboard with name, so can just paste into CIS search bar
	if (!(chk.Last~="[a-z]+")															; Check field values to see if proper demographics
		&& !(chk.First~="[a-z]+") 														; meaning names in ALL CAPS
		&& (chk.Acct~="\d{8}"))															; and EncNum present
	{
		MsgBox, 4132, Valid PDF, % ""
			. chk.Last ", " chk.First "`n"
			. "MRN " chk.MRN "`n"
			. "Acct " chk.Acct "`n"
			. "Ordering: " chk.Prov "`n"
			. "Study date: " chk.Date "`n`n"
			. "Is all the information correct?`n"
			. "If NO, reacquire demographics."
		IfMsgBox, Yes																; All tests valid
		{
			eventlog("Passed validation. Processing.")
			return																	; Select YES, return to processing Holter
		} 
		else 																		; Select NO, reacquire demographics
		{
			eventlog("Demographics valid. Wants to reacquire.")
			MsgBox, 4096, Adjust demographics, % chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
		}
	}
	else 																			; Not valid PDF, get demographics post hoc
	{
		eventlog("Demographics validation failed.")
		MsgBox, 4096,, % "Validation failed for:`n   " chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
	}
	; Either invalid PDF or want to correct values
	ptDem := Object()																; initialize/clear ptDem array
	ptDem["nameL"] := chk.Last															; Placeholder values for fetchGUI from PDF
	ptDem["nameF"] := chk.First
	ptDem["mrn"] := chk.MRN
	ptDem["DOB"] := chk.DOB
	ptDem["Sex"] := chk.Sex
	ptDem["Loc"] := chk.Loc
	ptDem["Account number"] := chk.Acct												; If want to force click, don't include Acct Num
	ptDem["Provider"] := trim(RegExReplace(RegExReplace(RegExReplace(chk.Prov,"i)^Dr(\.)?(\s)?"),"i)^[A-Z]\.(\s)?"),"(-MAIN| MD)"))
	ptDem["EncDate"] := chk.Date
	ptDem["Indication"] := chk.Ind
	
	fetchQuit:=false
	gosub fetchGUI
	gosub fetchDem
	/*	When fetchDem successfully completes,
	 *	replace the fields in demog with newly acquired values
	 */
	demog := RegExReplace(demog,"i)Last Name (.*)First Name","Last Name   " ptDem["nameL"] "`nFirst Name")
	demog := RegExReplace(demog,"i)First Name (.*)Middle Initial", "First Name   " ptDem["nameF"] "`nMiddle Initial")
	demog := RegExReplace(demog,"i)ID Number (.*)Date of Birth", "ID Number   " ptDem["mrn"] "`nDate of Birth")
	demog := RegExReplace(demog,"i)Date of Birth (.*)Sex", "Date of Birth   " ptDem["DOB"] "`nSex")
	demog := RegExReplace(demog,"i)Source (.*)Billing Code", "Source   " ptDem["Loc"] "`nBilling Code")
	demog := RegExReplace(demog,"i)Billing Code (.*)Recorder Format", "Billing Code   " ptDem["Account number"] "`nRecorder Format")
	demog := RegExReplace(demog,"i)Physician (.*)Scanned By", "Physician   " ptDem["Provider"] "`nScanned By")
	demog := RegExReplace(demog,"i)Test Date (.*)Analysis Date", "Test Date   " ptDem["EncDate"] "`nAnalysis Date")
	demog := RegExReplace(demog,"i)Reason for Test(.*)Group", "Reason for Test   " ptDem["Indication"] "`nGroup")	
	eventlog("Demog replaced.")
	
	return
}

CheckProcPR:
{
	eventlog("CheckProcPr")
	chk.Name := strVal(demog,"Name:","ID #:")												; Name
		chk.Last := trim(strX(chk.Name,"",1,1,",",1,1)," `r`n")									; NameL				must be [A-Z]
		chk.First := trim(strX(chk.Name,",",1,1,"",0)," `r`n")									; NameF				must be [A-Z]
	chk.MRN := strVal(demog,"ID #:","Second ID:")											; MRN
	chk.DOB := strVal(demog,"Date of Birth:","Age:")										; DOB
	chk.Sex := strVal(demog,"Sex:","Referring Physician:")									; Sex
	chk.Prov := strVal(demog,"Referring Physician:","Indications:")							; Ordering MD
	chk.Ind := strVal(demog,"Indications:","Medications:")									; Indication
	chk.Date := strVal(demog,"Date Recorded:","Date Processed:")							; Study date
	
	chkDT := parseDate(chk.Date)
	chkFilename := chk.MRN " * " chkDT.MM "-" chkDT.DD "-" chkDT.YYYY
	if FileExist(holterDir . "Archive\" . chkFilename . ".pdf") {
		FileDelete, %fileIn%
		eventlog(chk.MRN " PDF archive exists, deleting '" fileIn "'")
		fetchQuit := true
		return
	}
	
	Run , pdftotext.exe "%fileIn%" tempfull.txt,,min,wincons						; convert PDF all pages to txt file
	eventlog("Extracting full text.")	
	
	Clipboard := chk.Last ", " chk.First												; fill clipboard with name, so can just paste into CIS search bar
	if (!(chk.Last~="[a-z]+")															; Check field values to see if proper demographics
		&& !(chk.First~="[a-z]+") 														; meaning names in ALL CAPS
		&& (chk.Acct~="\d{8}"))															; and EncNum present
	{
		MsgBox, 4132, Valid PDF, % ""
			. chk.Last ", " chk.First "`n"
			. "MRN " chk.MRN "`n"
			. "Acct " chk.Acct "`n"
			. "Ordering: " chk.Prov "`n"
			. "Study date: " chk.Date "`n`n"
			. "Is all the information correct?`n"
			. "If NO, reacquire demographics."
		IfMsgBox, Yes																; All tests valid
		{
			eventlog("Demographics valid. Processing.")
			return																	; Select YES, return to processing Holter
		} 
		else 																		; Select NO, reacquire demographics
		{
			eventlog("Demographics valid. Wants to reacquire.")
			MsgBox, 4096, Adjust demographics, % chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
		}
	}
	else 																			; Not valid PDF, get demographics post hoc
	{
		eventlog("Demographics validation failed.")
		MsgBox, 4096,, % "Validation failed for:`n   " chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
	}
	; Either invalid PDF or want to correct values
	ptDem := Object()																; initialize/clear ptDem array
	ptDem["nameL"] := chk.Last															; Placeholder values for fetchGUI from PDF
	ptDem["nameF"] := chk.First
	ptDem["mrn"] := chk.MRN
	ptDem["DOB"] := chk.DOB
	ptDem["Sex"] := chk.Sex
	ptDem["Loc"] := chk.Loc
	ptDem["Account number"] := chk.Acct													; If want to force click, don't include Acct Num
	ptDem["Provider"] := trim(RegExReplace(RegExReplace(RegExReplace(chk.Prov,"i)^Dr(\.)?(\s)?"),"i)^[A-Z]\.(\s)?"),"(-MAIN| MD)"))
	ptDem["EncDate"] := chk.Date
	ptDem["Indication"] := chk.Ind
	
	fetchQuit:=false
	gosub fetchGUI
	gosub fetchDem
	/*	When fetchDem successfully completes,
	 *	replace the fields in demog with newly acquired values
	 */
	chk.Name := ptDem["nameL"] ", " ptDem["nameF"] 
	fldval["name_L"] := ptDem["nameL"]
	fldval["name_F"] := ptDem["nameF"]
	demog := RegExReplace(demog,"i`a)Name: (.*)\R","Name:   " chk.Name "   `n")
	demog := RegExReplace(demog,"i)ID #: (.*) Second ID:","ID #:   " ptDem["mrn"] "                   Second ID:")
	demog := RegExReplace(demog,"i)Date Of Birth: (.*) Age:", "Date Of Birth:   " ptDem["DOB"] "  Age:")
	demog := RegExReplace(demog,"i`a)Referring Physician: (.*)\R", "Referring Physician:   " ptDem["Provider"] "`n")
	demog := RegExReplace(demog,"i`a)Indications: (.*)\R", "Indications:   " ptDem["Indication"] "`n")	
	demog := RegExReplace(demog,"i`a)Date Recorded: (.*)\R", "Date Recorded:   " ptDem["EncDate"] "`n")
	demog := RegExReplace(demog,"i`a)Analyst: (.*) Hookup Tech:","Analyst:   $1 Hookup Tech:")
	demog := RegExReplace(demog,"i`a)Hookup Tech: (.*)\R","Hookup Tech:   $1   `n")
	demog .= "   Hookup time:   " ptDem["Hookup time"] "`n"
	demog .= "   Location:    " ptDem["Loc"] "`n"
	demog .= "   Acct Num:    " ptDem["Account number"] "`n"
	eventlog("Demog replaced.")
	
	return
}

Zio:
{
	eventlog("Holter_Zio")
	monType := "Zio"
	
	zcol := columns(newtxt,"","SIGNATURE",0,"Enrollment Period") ">>>end"
	demog := onecol(cleanblank(stregX(zcol,"\s+Date of Birth",1,0,"\s+(Supra)?ventricular tachycardia \(4",1)))
	
	gosub checkProcZio											; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	/* Holter PDF is valid. OK to process.
	 * Pulls text between field[n] and field[n+1], place in labels[n] name, with prefix "dem-" etc.
	 */
	
	znam := strVal(demog,"Name","Date of Birth")
	fieldColAdd("dem","Name_L",strX(znam, "", 1,0, ",", 1,1))
	fieldColAdd("dem","Name_F",strX(znam, ", ", 1,2, "", 0))
	
	fields[1] := ["Date of Birth","Prescribing Clinician","Patient ID","Managing Location","Gender","Primary Indication",">>>end"]
	labels[1] := ["DOB","Ordering","MRN","Site","Sex","Indication","end"]
	fieldvals(demog,1,"dem")
	
	tmp := columns(zcol,"\s+(Supra)?Ventricular","Preliminary Findings",0,"Ventricular")
	fieldColAdd("arr","SVT",scanfields(tmp,"Supraventricular Tachycardia \("))
	fieldColAdd("arr","VT",scanfields(tmp,"Ventricular Tachycardia \("))
	fieldColAdd("arr","Pauses",scanfields(tmp,"Pauses \("))
	fieldColAdd("arr","AVBlock",scanfields(tmp,"AV Block \("))
	fieldColAdd("arr","AF",scanfields(tmp,"Atrial Fibrillation"))
	
	znums := columns(zcol ">>>end","Enrollment Period",">>>end",1)
	
	fieldColAdd("time","Enrolled",chk.enroll)
	fieldColAdd("time","Analysis",chk.Analysis)
	
	zrate := columns(znums,"Heart Rate","Patient Events",1)
	fields[4] := ["Max ","Min ","Avg "]
	labels[4] := ["Max","Min","Avg"]
	fieldvals(zrate,4,"rate")
	
	zevent := columns(znums,"Number of Triggered Events:","Ectopics",1)
	fields[5] := ["Number of Triggered Events:","Findings within � 45 sec of Triggers:","Number of Diary Entries:","Findings within � 45 sec of Entries:"]
	labels[5] := ["Triggers","Trigger_Findings","Diary","Diary_Findings"]
	fieldvals(zevent,5,"event")
	
	zectopics := columns(znums,"Ectopics","Supraventricular Ectopy",1)
	fields[6] := ["Rare:","Occasional:","Frequent:"]
	labels[6] := ["Rare","Occ","Freq"]
	fieldvals(zectopics,6,"ecto")
	
	zsve := columns(znums,"Supraventricular Ectopy \(SVE/PACs\)","Ventricular Ectopy \(VE/PVCs\)",1)
	fields[7] := ["Isolated","Couplet","Triplet"]
	labels[7] := ["Single","Couplets","Triplets"]
	fieldvals(zsve,7,"sve")
	
	zve := columns(znums ">>>end","Ventricular Ectopy \(VE/PVCs\)",">>>end",1)
	fields[8] := ["Isolated","Couplet","Triplet","Longest Ventricular Bigeminy Episode","Longest Ventricular Trigeminy Episode"]
	labels[8] := ["Single","Couplets","Triplets","LongestBigem","LongestTrigem"]
	fieldvals(zve,8,"ve")
	
	zinterp := cleanspace(columns(newtxt,"Preliminary Findings","SIGNATURE",,"Final Interpretation"))
	zinterp := trim(StrX(zinterp,"",1,0,"Final Interpretation",1,20))
	fileout1 .= """INTERP""`n"
	fileout2 .= """" . zinterp . """`n"

return
}

CheckProcZio:
{
	tmp := trim(cleanSpace(stregX(zcol,"Report for",1,1,"Date of Birth",1)))
		chk.Last := trim(strX(tmp, "", 1,1, ",", 1,1))
		chk.First := trim(strX(tmp, ", ", 1,2, "", 0))
	chk.DOB := RegExReplace(strVal(demog,"Date of Birth","Prescribing Clinician"),"\s+\(.*(yrs|mos)\)")		; DOB
	chk.Prov:= strVal(demog,"Prescribing Clinician","Patient ID")												; Ordering MD
	chk.MRN := strVal(demog,"Patient ID","Managing Location")											; MRN
	chk.Loc := strVal(demog,"Managing Location","Gender")											; MRN
	chk.Sex := strVal(demog,"Gender","Primary Indication")											; Sex
	chk.Ind := RegExReplace(strVal(demog,"Primary Indication",">>>end"),"\(R00.0\)\s+")				; Indication
	
	demog := "Name   " chk.Last ", " chk.First "`n" demog
	
	tmp := oneCol(stregX(zcol,"Enrollment Period",1,0,"Heart\s+Rate",1))
		chk.enroll := strVal(tmp,"Enrollment Period","Analysis Time")
		chk.Date := strVal(chk.enroll,"hours",",")
		chk.Analysis := strVal(tmp,"Analysis Time","\(after")
		chk.enroll := stregX(chk.enroll,"",1,0,"   ",1)
	
	zcol := stregx(zcol,"\s+(Supra)?ventricular tachycardia \(",1,0,">>>end",1)
	
	/*	
	 *	Return from CheckProc for testing
	 */
		Return
	
	Clipboard := chk.Last ", " chk.First												; fill clipboard with name, so can just paste into CIS search bar
	if (!(chk.Last~="[a-z]+")															; Check field values to see if proper demographics
		&& !(chk.First~="[a-z]+") 														; meaning names in ALL CAPS
		&& (chk.Acct~="\d{8}"))															; and EncNum present
	{
		MsgBox, 4132, Valid PDF, % ""
			. chk.Last ", " chk.First "`n"
			. "MRN " chk.MRN "`n"
			. "Acct " chk.Acct "`n"
			. "Ordering: " chk.Prov "`n"
			. "Study date: " chk.Date "`n`n"
			. "Is all the information correct?`n"
			. "If NO, reacquire demographics."
		IfMsgBox, Yes																; All tests valid
		{
			return																	; Select YES, return to processing Holter
		} 
		else 																		; Select NO, reacquire demographics
		{
			MsgBox, 4096, Adjust demographics, % chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
		}
	}
	else 																			; Not valid PDF, get demographics post hoc
	{
		MsgBox, 4096,, % "Validation failed for:`n   " chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
	}
	; Either invalid PDF or want to correct values
	ptDem["nameL"] := chk.Last															; Placeholder values for fetchGUI from PDF
	ptDem["nameF"] := chk.First
	ptDem["mrn"] := chk.MRN
	ptDem["DOB"] := chk.DOB
	ptDem["Sex"] := chk.Sex
	ptDem["Loc"] := chk.Loc
	ptDem["Account number"] := chk.Acct													; If want to force click, don't include Acct Num
	ptDem["Provider"] := trim(RegExReplace(RegExReplace(RegExReplace(chk.Prov,"i)^Dr(\.)?(\s)?"),"i)^[A-Z]\.(\s)?"),"(-MAIN| MD)"))
	ptDem["EncDate"] := chk.Date
	ptDem["Indication"] := chk.Ind
	
	fetchQuit:=false
	gosub fetchGUI
	gosub fetchDem
	/*	When fetchDem successfully completes,
	 *	replace the fields in demog with newly acquired values
	 */
	chk.Name := ptDem["nameL"] ", " ptDem["nameF"] 
	demog := RegExReplace(demog,"i)Name(.*)Date of Birth","Name   " chk.Name "`nDate of Birth",,1)
	demog := RegExReplace(demog,"i)Date of Birth(.*)Prescribing Clinician","Date of Birth   " ptDem["DOB"] "`nPrescribing Clinician",,1)
	demog := RegExReplace(demog,"i)Prescribing Clinician(.*)Patient ID","Prescribing Clinician   " ptDem["Provider"] "`nPatient ID",,1)
	demog := RegExReplace(demog,"i)Patient ID(.*)Managing Location","Patient ID   " ptDem["MRN"] "`nManaging Location",,1)
	demog := RegExReplace(demog,"i)Managing Location(.*)Gender","Managing Location   " ptDem["Loc"] "`nGender",,1)
	demog := RegExReplace(demog,"i)Gender(.*)Primary Indication","Gender   " ptDem["Sex"] "`nPrimary Indication",,1)
	demog := RegExReplace(demog,"i)Primary Indication(.*)>>>end","Primary Indication   " ptDem["Indication"] "`n>>>end",,1)
	
	return
}

Event_LW:
{
	MsgBox, 16, File type error, Cannot process LifeWatch event recorders.`n`nPlease process this as a paper report.
	return
	
	fields := ["PATIENT INFORMATION","Name:","ID #:","4800 SAND POINT","DOB:","Sex:","Phone:"
		,"Monitor Type:","Diag:","Delivery Code:","Enrollment Period:","Date"
		,"SYMPTOMS:","ACTIVITY:","FINDINGS:","COMMENTS:","EVENT RECORDER DATA:"]
	n:=0
	Loop, parse, newtxt, `n,`r
	{
		i:=A_LoopField
		if !(i)													; skip entirely blank lines
			continue
		i = %i%
		newtxt .= i . "`n"						; strip left from right columns
	}
	for k, i in fields											; loop through result sections
	{
		j := fields[k+1]										; next field
		m := strx(newTxt,i,n,StrLen(i),j,1,StrLen(j)+1,n)		; extract between i and j
		m = %m%													; trim whitespace
		mm := cleancolon(m)
		m := mm
		blk.Insert(i)
		blk[i] := m												; associative array with result
		MsgBox,, % "(" k ")[" strlen(m) "] " i , % blk[i], 
	}
Return
}

Event_BGH:
{
	monType := "BGH"
	name := "Patient Name:   " trim(columns(newtxt,"Patient:","Enrollment Info",1,"")," `n")
	demog := columns(newtxt,"","Event Summary",,"Enrollment Info")
	enroll := RegExReplace(strX(demog,"Enrollment Info",1,0,"",0),": ",":   ")
	diag := "Diagnosis:   " trim(stRegX(demog,"`a)Diagnosis \(.*\R",1,1,"(Preventice)|(Enrollment Info)",1)," `n")
	demog := columns(demog,"\s+Patient ID","Diagnosis \(",,"Monitor   ") "#####"
	demog := columns(demog,"\s+Patient ID","#####",,"Gender","Date of Birth","Phone")		; columns get stuck in permanent loop
	demog := name "`n" demog "`n" diag "`n"
	
	demog0 := 
	Loop, parse, demog, `n,`r
	{
		i:=trim(A_LoopField)
		if !(i)													; skip entirely blank lines
			continue
		i = %i%
		demog0 .= i . "`n"						; strip left from right columns
	}
	demog := demog0
	
	gosub checkProcBGH											; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	fields[1] := ["Patient Name", "Patient ID", "Physician", "Gender", "Date of Birth", "Practice", "Diagnosis"]
	labels[1] := ["Name", "MRN", "Ordering", "Sex", "DOB", "VOID_Practice", "Indication"]
	fieldvals(demog,1,"dem")
	fldval["name_L"] := ptDem["nameL"]
	
	fields[2] := ["Period \(.*\)","Event Counts"]
	labels[2] := ["Test_date","VOID_Counts"]
	fieldvals(enroll,2,"dem")
	
	fields[3] := ["Critical","Total","Serious","Manual","Stable","Auto Trigger"]
	labels[3] := ["Critical","Total","Serious","Manual","Stable","Auto"]
	fieldvals(enroll,3,"counts")
	
	fileOut1 .= ",""Mon_type"""
	fileOut2 .= ",""Event"""

Return
}

strVal(hay,n1,n2,BO:="",ByRef N:="") {
/*	hay = search haystack
	n1	= needle1 begin string
	n2	= needle2 end string
	N	= return end position
*/
	;~ opt := "Oi" ((span) ? "s" : "") ")"
	opt := "Oi)"
	RegExMatch(hay,opt . n1 ":?(.*?)" n2 ":?",res,(BO)?BO:1)
	;~ MsgBox % trim(res[1]," `n") "`nPOS = " res.pos(1) "`nLEN = " res.len(1) "`n" res.value() "`n" res.len()
	N := res.pos()+res.len(1)

	return trim(res[1]," :`n")
}

CheckProcBGH:
{
	chk.Name := strVal(demog,"Patient Name","Patient ID")										; Name
		chk.First := trim(strX(chk.Name,"",1,1," ",1,1)," `r`n")									; NameL				must be [A-Z]
		chk.Last := trim(strX(chk.Name," ",0,1,"",0)," `r`n")										; NameF				must be [A-Z]
	chk.MRN := strVal(demog,"Patient ID","Physician")											; MRN
	chk.Prov := strVal(demog,"Physician","Gender")												; Ordering MD
	chk.Sex := strVal(demog,"Gender","Date of Birth")											; Sex
	chk.DOB := strVal(demog,"Date of Birth","Practice")											; DOB
	chk.Ind := strVal(demog,"Diagnosis",".*")													; Indication
	chk.Date := strVal(enroll,"Period \(.*\)","Event Counts")									; Study date
	
	Clipboard := chk.Last ", " chk.First												; fill clipboard with name, so can just paste into CIS search bar
	if (!(chk.Last~="[a-z]+")															; Check field values to see if proper demographics
		&& !(chk.First~="[a-z]+") 														; meaning names in ALL CAPS
		&& (chk.Acct~="\d{8}"))															; and EncNum present
	{
		MsgBox, 4132, Valid PDF, % ""
			. chk.Last ", " chk.First "`n"
			. "MRN " chk.MRN "`n"
			. "Acct " chk.Acct "`n"
			. "Ordering: " chk.Prov "`n"
			. "Study date: " chk.Date "`n`n"
			. "Is all the information correct?`n"
			. "If NO, reacquire demographics."
		IfMsgBox, Yes																; All tests valid
		{
			return																	; Select YES, return to processing Holter
		} 
		else 																		; Select NO, reacquire demographics
		{
			MsgBox, 4096, Adjust demographics, % chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
		}
	}
	else 																			; Not valid PDF, get demographics post hoc
	{
		MsgBox, 4096,, % "Validation failed for:`n   " chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
	}
	; Either invalid PDF or want to correct values
	ptDem["nameL"] := chk.Last															; Placeholder values for fetchGUI from PDF
	ptDem["nameF"] := chk.First
	ptDem["mrn"] := chk.MRN
	ptDem["DOB"] := chk.DOB
	ptDem["Sex"] := chk.Sex
	ptDem["Loc"] := chk.Loc
	ptDem["Account number"] := chk.Acct													; If want to force click, don't include Acct Num
	ptDem["Provider"] := trim(RegExReplace(RegExReplace(RegExReplace(chk.Prov,"i)^Dr(\.)?(\s)?"),"i)^[A-Z]\.(\s)?"),"(-MAIN| MD)"))
	ptDem["EncDate"] := chk.Date
	ptDem["Indication"] := chk.Ind
	
	fetchQuit:=false
	gosub fetchGUI
	gosub fetchDem
	/*	When fetchDem successfully completes,
	 *	replace the fields in demog with newly acquired values
	 */
	chk.Name := ptDem["nameF"] " " ptDem["nameL"] 
		fldval["name_L"] := ptDem["nameL"]
		fldval["name_F"] := ptDem["nameF"]
	demog := RegExReplace(demog,"i)Patient Name: (.*)Patient ID","Patient Name:   " chk.Name "`nPatient ID")
	demog := RegExReplace(demog,"i)Patient ID(.*)Physician","Patient ID   " ptDem["mrn"] "`nPhysician")
	demog := RegExReplace(demog,"i)Physician(.*)Gender", "Physician   " ptDem["Provider"] "`nGender")
	demog := RegExReplace(demog,"i)Gender(.*)Date of Birth", "Gender   " ptDem["Sex"] "`nDate of Birth")
	demog := RegExReplace(demog,"i)Date of Birth(.*)Practice", "Date of Birth   " ptDem["DOB"] "`nPractice")	
	enroll := RegExReplace(enroll,"i)Date Recorded: (.*)\R", "Date Recorded:   " ptDem["EncDate"] "`n")
	;~ demog := RegExReplace(demog,"i`a)Analyst: (.*) Hookup Tech:","Analyst:   $1 Hookup Tech:")
	;~ demog := RegExReplace(demog,"i`a)Hookup Tech: (.*)\R","Hookup Tech:   $1   `n")
	
	return
}

oneCol(txt) {
/*	Break text block into a single column 
	based on logical break points in title (first) row
*/
	lastpos := 1
	Loop																		; Iterate each column
	{
		Loop, parse, txt, `n,`r													; Read through text block
		{
			i := A_LoopField
			
			if (A_index=1) {
				pos := RegExMatch(i	"  "										; Add "  " to end of scan string
								,"O)(?<=(\s{2}))[^\s]"							; Search "  text" as each column 
								,col
								,lastpos+1)										; search position to find next "  "
				
				if !(pos) {														; no match beyond, have hit max column
					max := true
				}
			}
			
			len := (max) ? strlen(i) : pos-lastpos								; length of string to return (max gets to end of line)
			
			str := substr(i,lastpos,len)										; string to return
			
			result .= str "`n"													; add to result
			;~ MsgBox % result
		}
		if !(pos) {																; break out if at max column
			break
		}
		lastpos := pos															; set next start point
	}
	return result . ">>>end"
}

columns(x,blk1,blk2,incl:="",col2:="",col3:="",col4:="") {
/*	Returns string as a single column.
	x 		= input string
	blk1	= leading string to start block
	blk2	= ending string to end block
	incl	= if null, include blk1 string; if !null, remove blk1 string
	col2	= string demarcates start of COLUMN 2
	col3	= string demarcates start of COLUMN 3
	col4	= string demarcates start of COLUMN 4
*/
	blk1 := rxFix(blk1,"O",1)													; Adds "O)" to blk1
	blk2 := rxFix(blk2,"O",1)
	RegExMatch(x,blk1,blo1)														; Creates blo1 object out of blk1 match in x
	RegExMatch(x,blk2,blo2)
	
	txt := stRegX(x,blk1,1,((incl) ? 1 : 0),blk2,1)
	;~ MsgBox % txt
	col2 := RegExReplace(col2,"\s+","\s+")
	col3 := RegExReplace(col3,"\s+","\s+")
	col4 := RegExReplace(col4,"\s+","\s+")
	
	loop, parse, txt, `n,`r										; find position of columns 2, 3, and 4
	{
		i:=A_LoopField
		if (t:=RegExMatch(i,col2))
			pos2:=t
		if (t:=RegExMatch(i,col3))
			pos3:=t
		if (t:=RegExMatch(i,col4))
			pos4:=t
	}
	loop, parse, txt, `n,`r
	{
		i:=A_LoopField
		txt1 .= substr(i,1,pos2-1) . "`n"
		if (col4) {
			pos4ck := pos4
			while !(substr(i,pos4ck-1,1)=" ") {
				pos4ck := pos4ck-1
			}
			txt4 .= substr(i,pos4ck) . "`n"
			txt3 .= substr(i,pos3,pos4ck-pos3) . "`n"
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"
			continue
		} 
		if (col3) {
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"
			txt3 .= substr(i,pos3) . "`n"
			continue
		}
		txt2 .= substr(i,pos2) . "`n"
	}
	return txt1 . txt2 . txt3 . txt4
}

scanfields(x,lbl) {
/*	Scans text for block from lbl to next lbl
*/
	i := trim(stregX(x,"[\r\n]+" lbl,1,0,"[\r\n]+\w",1)," `r`n")
	if instr(i,"Episodes") {
		i := trim(columns(i ">>>end","",">>>end",1,"Episodes")," `r`n")
	}
	i := RegExReplace(i,"i)None found","0")
	j := cleanblank(substr(i,(i~="[\r\n]+")))
	return j
}

fieldvals(x,bl,bl2) {
/*	Matches field values and results. Gets text between FIELDS[k] to FIELDS[k+1]. Excess whitespace removed. Returns results in array BLK[].
	x	= input text
	bl	= which FIELD number to use
	bl2	= label prefix
*/
	global fields, labels, fldval
	
	for k, i in fields[bl]
	{
		pre := bl2
		j := fields[bl][k+1]
		m := (j) ?	strVal(x,i,j,n,n)			;trim(stRegX(x,i,n,1,j,1,n), " `n")
				:	trim(strX(SubStr(x,n),":",1,1,"",0)," `n")
		lbl := labels[bl][A_index]
		if (lbl~="^\w{3}:") {											; has prefix e.g. "dem:"
			pre := substr(lbl,1,3)
			lbl := substr(lbl,5)
		}
		cleanSpace(m)
		cleanColon(m)
		fldval[lbl] := m
		
		formatField(pre,lbl,m)
	}
}

/*	rxFix
	in	= input string, may or may or not include "Oim)" option modifiers
	req	= required modifiers to output
	spc	= replace spaces
*/
rxFix(hay,req,spc:="")
{
	opts:="^[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"
	out := (hay~=opts) ? req . hay : req ")" hay
	if (spc) {
		out := RegExReplace(out,"\s+","\s+")
	}
	return out
}

/* StrX parameters
StrX( H, BS,BO,BT, ES,EO,ET, NextOffset )

Parameters:
H = HayStack. The "Source Text"
BS = BeginStr. 
Pass a String that will result at the left extreme of Resultant String.
BO = BeginOffset. 
Number of Characters to omit from the left extreme of "Source Text" while searching for BeginStr
Pass a 0 to search in reverse ( from right-to-left ) in "Source Text"
If you intend to call StrX() from a Loop, pass the same variable used as 8th Parameter, which will simplify the parsing process.
BT = BeginTrim.
Number of characters to trim on the left extreme of Resultant String
Pass the String length of BeginStr if you want to omit it from Resultant String
Pass a Negative value if you want to expand the left extreme of Resultant String
ES = EndStr. Pass a String that will result at the right extreme of Resultant String
EO = EndOffset.
Can be only True or False.
If False, EndStr will be searched from the end of Source Text.
If True, search will be conducted from the search result offset of BeginStr or from offset 1 whichever is applicable.
ET = EndTrim.
Number of characters to trim on the right extreme of Resultant String
Pass the String length of EndStr if you want to omit it from Resultant String
Pass a Negative value if you want to expand the right extreme of Resultant String
NextOffset : A name of ByRef Variable that will be updated by StrX() with the current offset, You may pass the same variable as Parameter 3, to simplify data parsing in a loop
*/
StrX( H,  BS="",BO=0,BT=1,   ES="",EO=0,ET=1,  ByRef N="" ) { ;    | by Skan | 19-Nov-2009
Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
<0)?(1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
+(0-ET)):(X+P)):(X)))-P) ; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
}

stRegX_old(h,BS="",BO=1,BT=0, ES="",ET=0, ByRef N="") {
/*	modified version: searches from BS to "   "
	h = Haystack
	BS = beginning string
	BO = beginning offset
	BT = beginning trim, TRUE or FALSE
	ES = ending string
	ET = ending trim, TRUE or FALSE
	N = variable for next offset
*/
	BS .= "(.*?)\s{3}"
	rem:="[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"
	pos0 := RegExMatch(h,((BS~=rem)?"Oim"BS:"Oim)"BS),bPat,((BO<1)?1:BO))
	pos1 := RegExMatch(h,((ES~=rem)?"Oim"ES:"Oim)"ES),ePat,pos0+bPat.len)
	N := pos1+((ET)?0:(ePat.len))
	return substr(h,pos0+((BT)?(bPat.len):0),N-pos0-bPat.len)
}

stRegX(h,BS="",BO=1,BT=0, ES="",ET=0, ByRef N="") {
/*	modified version: searches from BS to "   "
	h = Haystack
	BS = beginning string
	BO = beginning offset
	BT = beginning trim, TRUE or FALSE
	ES = ending string
	ET = ending trim, TRUE or FALSE
	N = variable for next offset
*/
	;~ BS .= "(.*?)\s{3}"
	rem:="^[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"										; All the possible regexmatch options
	
	pos0 := RegExMatch(h,((BS~=rem)?"Oim"BS:"Oim)"BS),bPat,((BO)?BO:1))
	/*	Ensure that BS begins with at least "Oim)" to return [O]utput, case [i]nsensitive, and [m]ultiline searching
		Return result in "bPat" (beginning pattern) object
		If (BO), start at position BO, else start at 1
	*/
	pos1 := RegExMatch(h,((ES~=rem)?"Oim"ES:"Oim)"ES),ePat,pos0+bPat.len())
	/*	Ensure that ES begins with at least "Oim)"
		Resturn result in "ePat" (ending pattern) object
		Begin search after bPat result (pos0+bPat.len())
	*/
	bmod := (BT) ? bPat.len() : 0
	emod := (ET) ? 0 : ePat.len()
	N := pos1+emod
	/*	Final position is start of ePat match + modifier
		If (ET), add nothing, else add ePat.len()
	*/
	return substr(h,pos0+bmod,(pos1+emod)-(pos0+bmod))
	/*	Start at pos0
		If (BT), add bPat.len(), else stay at pos0 (will include BS in result)
		substr length is position of N (either pos1 or include ePat) less starting pos0
	*/
}

formatField(pre, lab, txt) {
	global monType, Docs, ptDem
	if (txt ~= "\d{1,2} hr \d{1,2} min") {
		StringReplace, txt, txt, %A_Space%hr%A_space% , :
		StringReplace, txt, txt, %A_Space%min , 
	}
	txt:=RegExReplace(txt,"i)BPM|Event(s)?|Beat(s)?|( sec(s)?)")			; 	Remove units from numbers
	txt:=RegExReplace(txt,"(:\d{2}?)(AM|PM)","$1 $2")						;	Fix time strings without space before AM|PM
	txt := trim(txt)
	
	if (lab="Ordering") {
		tmpCrd := checkCrd(RegExReplace(txt,"i)^Dr(\.)?\s"))
		fieldColAdd(pre,lab,tmpCrd.best)
		fieldColAdd(pre,lab "_grp",tmpCrd.group)
		fieldColAdd(pre,lab "_eml",Docs[tmpCrd.Group ".eml",ObjHasValue(Docs[tmpCrd.Group],tmpCrd.best)])
		return
	}
	
;	Lifewatch Holter specific search fixes
	if (monType="H") {
		if txt ~= ("^[0-9]+.*at.*(AM|PM)$") {								;	Split timed results "139 at 8:31:47 AM" into two fields
			tx1 := trim(strX(txt,,1,1," at",1,3))							;		labels e.g. xxx and xxx_time
			tx2 := trim(strX(txt," at",1,3,"",1,0))							;		result e.g. "139" and "8:31:47 AM"
			fieldColAdd(pre,lab,tx1)
			fieldColAdd(pre,lab "_time",tx2)
			return
		}
		if (lab~="i)(Longest|Fastest)") {
			fieldColAdd(pre,lab,txt)
			fieldColAdd(pre,lab "_time","")
			return
		}
		if (txt ~= "^[0-9]+\s\([0-9.]+\%\)$") {								;	Split percents |\(.*%\)
			tx1 := trim(strX(txt,,1,1,"(",1,1))
			tx2 := trim(strX(txt,"(",1,1,"%",1,0))
			fieldColAdd(pre,lab,tx1)
			fieldColAdd(pre,lab "_per",tx2)
			return
		}
		if (txt ~= "^[0-9,]{1,}\/[0-9,]{1,}$") {							;	Split multiple number value results "5/0" into two fields, ignore date formats (5/1/12)
			tx1 := strX(txt,,1,1,"/",1,1,n)
			tx2 := SubStr(txt,n+1)
			lb1 := strX(lab,,1,1,"_",1,1,n)									;	label[] fields are named "xxx_yyy", split into "xxx" and "yyy"
			lb2 := SubStr(lab,n+1)
			fieldColAdd(pre,lb1,tx1)
			fieldColAdd(pre,lb2,tx2)
			return
		}
	}
	
;	Preventice Holter specific fixes
	if (monType="PR") {
		if (lab="Name") {
			fieldColAdd(pre,"Name_L",trim(strX(txt,"",1,0,",",1,1)))
			fieldColAdd(pre,"Name_F",trim(strX(txt,",",1,1,"",0)))
			return
		}
		if (RegExMatch(txt,"O)^(\d{1,2})\s+hr,\s+(\d{1,2})\s+min",tx)) {
			fieldColAdd(pre,lab,zDigit(tx.value(1)) ":" zDigit(tx.value(2)))
			return
		}
		if (RegExMatch(txt,"O)^([0-9.]+).*at.*(\d{2}:\d{2}:\d{2})(AM|PM)?$",tx)) {		;	Split timed results "139 at 8:31:47 AM" into two fields
			fieldColAdd(pre,lab,tx.value(1))
			fieldColAdd(pre,lab "_time",tx.value(2))
			return
		}
	}

;	Body Guardian Heart specific fixes
	if (monType="BGH") {
		if (lab="Name") {
			ptDem["nameL"] := strX(txt," ",0,1,"",0)
			ptDem["nameF"] := strX(txt,"",1,0," ",1,1)
			fieldColAdd(pre,"Name_L",ptDem["nameL"])
			fieldColAdd(pre,"Name_F",ptDem["nameF"])
			return
		}
		if (lab="Test_date") {
			RegExMatch(txt,"O)(\d{1,2}/\d{1,2}/\d{4}).* (\d{1,2}/\d{1,2}/\d{4})",dt)
			fieldColAdd(pre,lab,dt.value(1))
			fieldColAdd(pre,lab "_end",dt.value(2))
			return
		}
	}
	
;	ZIO patch specific search fixes
	if (monType="Zio") {
		if (RegExMatch(txt,"(\d){1,2} days (\d){1,2} hours ",tmp)) {		;	Split recorded/analyzed time in to Days and Hours
			fieldColAdd(pre,lab "_D",strX(tmp,"",1,1, " days",1,5))
			fieldColAdd(pre,lab "_H",strX(tmp," days",1,6, " hours",1,6))
			fieldColAdd(pre,lab "_Dates",substr(txt,instr(txt," hours ")+7))
			return
		}
		if InStr(txt,"(at ") {												;	Split timed results "139 (at 8:31:47 AM)" into two fields
			tx1 := strX(txt,,1,1,"(at ",1,4,n)
			tx2 := trim(SubStr(txt,n+4), " )")
			fieldColAdd(pre,lab,tx1)
			fieldColAdd(pre,lab "_time",tx2)
			return
		}
		if (RegExMatch(txt,"i)[a-z]+\s+[\>\<\.0-9%]+\s+\d",tmp)) {			;	Split "RARE <1.0% 2457" into result "2457" and text quant "RARE <1.0%"
			tx1 := substr(txt,1,StrLen(tmp)-2)
			tx2 := substr(txt,strlen(tmp))
			fieldColAdd(pre,lab,tx2)
			fieldColAdd(pre,lab "_amt",tx1)
			return
		}
		if (txt ~= "3rd.*\)") {												;	fix AV block field
			txt := substr(txt, InStr(txt, ")")+2)
		}
		if (txt=="None found") {											;	fix 0 results
			txt := "0"
		}
	}
	
	fieldColAdd(pre,lab,txt)
	return
}

fieldColAdd(pre,lab,txt) {
	global fileOut1, fileOut2
	fileOut1 .= """" pre "-" lab ""","
	fileOut2 .= """" txt ""","
	return
}

checkCrd(x) {
/*	Compares pl_ProvCard vs array of cardiologists
	x = name
	returns array[match score, best match, best match group]
*/
	global Docs
	fuzz := 0.1
	for rowidx,row in Docs
	{
		if (substr(rowIdx,-3)=".eml")
			continue
		for colidx,item in row
		{
			res := fuzzysearch(x,item)
			if (res<fuzz) {
				fuzz := res
				best:=item
				group:=rowidx
			}
		}
	}
	return {"fuzz":fuzz,"best":best,"group":group}
}

httpComm(verb) {
	; consider two parameters?
	;~ global servFold
	servfold := "patlist"
	whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")							; initialize http request in object whr
		whr.Open("GET"															; set the http verb to GET file "change"
			, "https://depts.washington.edu/pedcards/change/direct.php?" 
				. ((servFold="testlist") ? "test=true&" : "") 
				. "do=" . verb
			, true)
		whr.Send()																; SEND the command to the address
		whr.WaitForResponse()	
	return whr.ResponseText													; the http response
}

cleancolon(ByRef txt) {
	if substr(txt,1,1)=":" {
		txt:=substr(txt,2)
		txt = %txt%
	}
	return txt
}

cleanspace(ByRef txt) {
	StringReplace txt,txt,`n,%A_Space%, All
	StringReplace txt,txt,%A_Space%.%A_Space%,.%A_Space%, All
	loop
	{
		StringReplace txt,txt,%A_Space%%A_Space%,%A_Space%, UseErrorLevel
		if ErrorLevel = 0	
			break
	}
	return txt
}

cleanblank(txt) {
	Loop, parse, txt, `n,`r                            ; clean up maintxt 
	{
		i:=A_LoopField
		if !(trim(i))                                    ; skip entirely blank lines 
		  continue
		newTxt .= i . "`n"                              ; only add lines with text in it 
	} 
	return newTxt
}

ObjHasValue(aObj, aValue, rx:="") {
; modified from http://www.autohotkey.com/board/topic/84006-ahk-l-containshasvalue-method/	
    for key, val in aObj
		if (rx="RX") {
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

eventlog(event) {
	global user
	comp := A_ComputerName
	FormatTime, sessdate, A_Now, yyyy.MM
	FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
	name := "logs/" . sessdate . ".log"
	txt := now " [" user "/" comp "] " event "`n"
	filePrepend(txt,name)
;	FileAppend, % timenow " ["  user "/" comp "] " event "`n", % "logs/" . sessdate . ".log"
}

FilePrepend( Text, Filename ) { 
/*	from haichen http://www.autohotkey.com/board/topic/80342-fileprependa-insert-text-at-begin-of-file-ansi-text/?p=510640
*/
    file:= FileOpen(Filename, "rw")
    text .= File.Read()
    file.pos:=0
    File.Write(text)
    File.Close()
}

parseDate(x) {
; Disassembles "2/9/2015" or "2/9/2015 8:31" into Yr=2015 Mo=02 Da=09 Hr=08 Min=31
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	if (x~="i)(\d{1,2})[\-\s\.](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\-\s\.](\d{2,4})") {		; 03 Jan 2016
		StringSplit, DT, x, %A_Space%-.
		return {"DD":zDigit(DT1), "MM":zDigit(objHasValue(mo,DT2)), "MMM":DT2, "YYYY":year4dig(DT3)}
	}
	if (x~="\d{1,2}_\d{1,2}_\d{2,4}") {											; 03_06_17 or 03_06_2017
		StringSplit, DT, x, _
		return {"MM":zDigit(DT1), "DD":zDigit(DT2), "MMM":mo[DT2], "YYYY":year4dig(DT3)}
	}
	if (x~="\d{4}-\d{2}-\d{2}") {												; 2017-02-11
		StringSplit, DT, x, -
		return {"YYYY":DT1, "MM":DT2, "DD":DT3}
	}
	if (x~="i)^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2}, \d{4}") {			; Mar 9, 2015 (8:33 am)?
		StringSplit, DT, x, %A_Space%
		StringSplit, DHM, DT4, :
		return {"MM":zDigit(objHasValue(mo,DT1)),"DD":zDigit(trim(DT2,",")),"YYYY":DT3
			,	hr:zDigit((DT5~="i)p")?(DHM1+12):DHM1),min:DHM2}
	}
	StringSplit, DT, x, %A_Space%
	StringSplit, DY, DT1, /
	StringSplit, DHM, DT2, :
	return {"MM":zDigit(DY1), "DD":zDigit(DY2), "YYYY":year4dig(DY3), "hr":zDigit(DHM1), "min":zDigit(DHM2), "Date":DT1, "Time":DT2}
}

niceDate(x) {
	if !(x)
		return error
	FormatTime, x, %x%, MM/dd/yyyy
	return x
}

year4dig(x) {
	if (StrLen(x)=4) {
		return x
	}
	if (StrLen(x)=2) {
		return (x<50)?("20" x):("19" x)
	}
	return error
}

zDigit(x) {
; Add leading zero to a number
	return SubStr("0" . x, -1)
}

#Include CMsgBox.ahk
#Include xml.ahk
#Include sift3.ahk
