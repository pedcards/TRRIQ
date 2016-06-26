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

SplitPath, A_ScriptDir,,fileDir
IfInString, fileDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	isAdmin := true
	holterDir := ".\Holter PDFs\"
	importFld := ".\Import\"
	chipDir := ".\Chipotle\"
} else {
	isAdmin := false
	holterDir := "..\Holter PDFs\"
	importFld := "..\Import\"
	chipDir := "\\childrens\files\HCChipotle\"
}
user := A_UserName

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
	if !(instr(tmp4,"seattlechildrens.org")) {						; skip non-children's providers
		continue
	}																; Otherwise format Crd name to first initial, last name
	tmpIdx += 1
	StringSplit, tmpPrv, tmp1, %A_Space%`"
	;tmpPrv := substr(tmpPrv1,1,1) . ". " . tmpPrv2					; F. Last
	tmpPrv := tmpPrv2 ", " tmpPrv1									; Last, First
	Docs[tmpGrp,tmpIdx]:=tmpPrv
	Docs[tmpGrp ".eml",tmpIdx] := tmp4
}

y := new XML(chipDir "currlist.xml")

siteVals := {"CRD":"Seattle","EKG":"EKG lab","ECO":"ECHO lab","CRDBCSC":"Bellevue","CRDEVT":"Everett","CRDTAC":"Tacoma","CRDTRI":"Tri Cities","CRDWEN":"Wenatchee","CRDYAK":"Yakima"}
demVals := ["MRN","Account Number","DOB","Sex","Loc","Provider"]						; valid field names for parseClip()

if !(phase) {
	phase := CMsgBox("Which task?","","*&Upload new Holter|&Process PDF","Q","")
}
if (instr(phase,"new")) {
	Loop 
	{
		ptDem := Object()
		gosub fetchGUI								; Draw input GUI
		gosub fetchDem								; Grab demographics from CIS until accept
		gosub zybitSet								; Fill in Zybit demographics
	}
	ExitApp
}
if (instr(phase,"PDF")) {
	holterLoops := 0								; Reset counters
	holtersDone := 
	loop, %holterDir%*.pdf							; Process all PDF files in holterDir
	{
		fileNam := RegExReplace(A_LoopFileName,"i)\.pdf")				; fileNam is name only without extension, no path
		fileIn := A_LoopFileFullPath									; fileIn has complete path \\childrens\files\HCCardiologyFiles\EP\HoltER Database\Holter PDFs\steve.pdf
		FileGetTime, fileDt, %fileIn%, C								; fildDt is creatdate/time 
		if (substr(fileDt,-5)="000000") {								; skip files with creation TIME midnight (already processed)
			continue
		}
		gosub MainLoop													; process the PDF
		if (fetchQuit=true) {											; [x] out of fetchDem means skip this file
			continue
		}
		holterLoops++													; increment counter for processed counter
		holtersDone .= A_LoopFileName "->" filenameOut ".pdf`n"			; add to report
	}
	MsgBox,, % "Holters processed (" holterLoops ")", % holtersDone
	ExitApp
	/* Consider asking if complete. The MA's appear to run one PDF at a time, despite the efficiency loss.
	*/
}

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
					if (clk.value) {													; extract provider.value to LAST,FIRST (strip MD, PHD, MI, etc)
						tmp := strX(clk.value,,1,0, ",",1,1) ", " strX(clk.value,",",1,2, " ",1,1)
					}
					if (ptDem.Provider) {												; Provider already exists
						MsgBox, 4148
							, Provider already exists
							, % "Replace " ptDem.Provider "`n with `n" tmp "?"
						IfMsgBox, Yes													; Check before replacing
						{
							ptDem.Provider := tmp
						}
					} else {															; Otherwise populate ptDem.Provider
						ptDem.Provider := tmp
					}
					mdX[4] := mouseXpos													; demographics grid[4,1]
					mdY[1] := mouseYpos
					mdProv := true														; we have got Provider
					WinGetTitle, mdTitle, ahk_id %mouseWinID%
					gosub getDemName													; extract patient name, MRN from window title 
					
				}																		;(this is why it must be sister or parent VM).
				if (clk.field = "Account Number") {
					ptDem["Account Number"] := clk.value
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
					ptDem["MRN"] := mouseGrab(mdX[1],mdY[2])							; grab remaining demographic values
					ptDem["DOB"] := mouseGrab(mdX[2],mdY[2])
					ptDem["Sex"] := substr(mouseGrab(mdX[3],mdY[1]),1,1)
					tmp := mouseGrab(mdX[3],mdY[3])										; grab Encounter Type field
						tmpType := strX(tmp,,1,0, " [",1,2)								; Type is everything up to " ["
						tmpDate := strX(tmp," [",1,2, " ",1,1)							; Date is anything between " [" and " "
					ptDem["Type"] := tmpType
					
					if (ptDem.Type="Outpatient") {
						ptDem["Loc"] := mouseGrab(mdX[3]+mdXd*0.5,mdY[2])				; most outpatient locations are short strings, click the right half of cell to grab location name
						ptDem["EncDate"] := tmpDate
					}
					if (ptDem.Type="Inpatient") {										; could be actual inpatient or in SurgCntr
						ptDem["Loc"] := "Inpatient"										; date is date of admission
						ptDem["EncDate"] := tmpDate
					}
					if (ptDem.Type="Day Surg") {
						ptDem["Loc"] := "SurgCntr"										; fill the ptDem.Loc field
						ptDem["EncDate"] := tmpDate										; date in SurgCntr
					}
					mdProv := false														; processed demographic fields,
					mdAcct := false														; so reset check bits
				}
				if !(clk.field~="(Provider|Account Number)") {							; all other values
					ptDem[clk.field] := (clk.value) ? clk.value : ptDem[clk.field]		; populate ptDem.field with value; if value=null, keep same]
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
	BlockInput, On
	MouseMove, %x%, %y%, 0
	Click 2
	BlockInput, Off
	sleep 100
	ClipWait
	clk := parseClip(clipboard)
	return clk.value
}

parseClip(clip) {
	global demVals
	StringSplit, val, clip, :															; break field into val1:val2
	if (ObjHasValue(demVals, val1)) {													; field name in demVals, e.g. "MRN","Account Number","DOB","Sex","Loc","Provider"
		return {"field":val1, "value":val2}
	}
	if (clip~="Outpatient\s\[") {														; Outpatient type
		return {"field":"Type", "value":clip}											; return original clip string (to be broken later)
	}
	if (clip~="Inpatient\s\[") {														; Inpatient types
		return {"field":"Type", "value":"Inpatient"}									; return "Inpt"
	}
	if (clip~="Day Surg.*\s\[") {														; Day Surg type
		return {"field":"Type"
				, "value":"Day Surg"													; return "Day Surg"
				, "date":strX(clip," [",1,2, " ",1,1)}									; and date
	}
	if (clip~="Emergency") {															; Emergency type
		return {"field":"Type"
				, "value":"Emergency"													; return "Day Surg"
				, "date":strX(clip," [",1,2, " ",1,1)}									; and date
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
	fYd := 30,	fXd := 80									; fetchGUI delta Y, X
	fX1 := 12,	fX2 := fX1+fXd								; x pos for title and input fields
	fW1 := 60,	fW2 := 190									; width for title and input fields
	fH := 20												; line heights
	fY := 10												; y pos to start
	EncNum := ptDem["Account Number"]						; we need these non-array variables for the Gui statements
	encDT := parseDate(ptDem.EncDate).YYYY . parseDate(ptDem.EncDate).MM . parseDate(ptDem.EncDate).DD
	fTxt := "	To auto-grab demographic info:`n"
		.	"		1) Double-click Account Number #`n"
		.	"		2) Double-click Provider"
	Gui, fetch:Destroy
	Gui, fetch:+AlwaysOnTop
	Gui, fetch:Add, Text, % "x" fX1 , % fTxt	
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd*2) " w" fW1 " h" fH , First
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["nameF"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Last
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["nameL"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , MRN
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["MRN"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , DOB
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH, % ptDem["DOB"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Date placed
	Gui, fetch:Add, DateTime, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " vEncDt CHOOSE" encDT, MM/dd/yyyy
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Encounter #
	Gui, fetch:Add, Edit, % "x" fX2 " y" fY-4 " w" fW2 " h" fH " vEncNum", % encNum
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Ordering MD
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["Provider"]
	Gui, fetch:Add, Button, % "x" fX1+10 " y" (fY += fYD) " h" fH+10 " w" fW1+fW2 " gfetchSubmit", Submit!
	Gui, fetch:Show, AutoSize, Enter Demographics
	return
}

fetchGuiClose:
	Gui, fetch:destroy
	getDem := false																	; break out of fetchDem loop
	fetchQuit := true
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
	if (ptDem.Type~="i)(Inpatient|Emergency)") {								; Inpt & ER, we must find who recommended it from the Chipotle schedule
		gosub assignMD
	} else if (ptDem.Loc~="i)SurgCntr") {										; SURGCNTR, find who recommended it
		gosub getMD
	} else if (ptDem.Loc~="i)(EKG|ECO|DCT)") {									; Any outpatient EKG ECO DCT account (Holter-only), ask for ordering MD
		gosub getMD
;	} else if !(ptDem.Loc~="i)(CRD|EKG|ECO|DCT|SurgCntr).*") {					; Not any CRDxxx location, must be an appropriate encounter (CRD,EKG,ECO,DCT or Inpt or ER)
	} else {																	; otherwise fail
		MsgBox % "Invalid Loc`n" ptDem.Loc
		gosub fetchGUI
		return
	}
	if !(ptDem.Provider) {
		gosub getMD																; No CRD provider, ask for it.
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
		gosub indGUI
		WinWaitClose, Enter indications
		if (indChoices)															; loop until we have filled indChoices
			break
	}
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
	Gui, ind:Add, Text, , % "Enter indications: " ptDem["Account Number"] " - " encNum
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
	RunWait, pdftotext.exe -l 2 -table -fixed 3 "%fileIn%" temp.txt					; convert PDF to txt file
	newTxt:=""																		; clear the full txt variable
	FileRead, maintxt, temp.txt														; load into maintxt
	Loop, parse, maintxt, `n,`r														; clean up maintxt
	{					
		i:=A_LoopField					
		if !(i)																		; skip entirely blank lines
			continue					
		newTxt .= i . "`n"															; only add lines with text in it
	}					
	FileDelete tempfile.txt															; remove any leftover tempfile
	FileAppend %newtxt%, tempfile.txt												; create new tempfile with newtxt result
	FileMove tempfile.txt, .\tempfiles\%fileNam%.txt								; move a copy into tempfiles for troubleshooting

	blocks := Object()																; clear all objects
	fields := Object()
	fldval := {}
	labels := Object()
	blk := Object()
	blk2 := Object()
	fileOut1 := fileOut2 := ""
	summBl := summ := ""
	
	if (InStr(maintxt,"Holter")) {															; Processing loop based on identifying string in maintxt
		gosub Holter
	} else if (InStr(maintxt,"TRANSTELEPHONIC ARRHYTHMIA")) {
		gosub EventRec
	} else if (RegExMatch(maintxt,"i)zio.*xt.*patch")) {
		gosub Zio
	} else {
		MsgBox No match!
		ExitApp
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
	FileDelete, %importFld%%fileNameOut%.csv												; clear any previous CSV
	FileAppend, %fileOut%, %importFld%%fileNameOut%.csv										; create a new CSV
	FileCopy, %importFld%%fileNameOut%.csv, .\tempfiles\*.*, 1								; create a copy of CSV in tempfiles
	FileMove, %fileIn%, %holterDir%%filenameOut%.pdf, 1										; move the PDF to holterDir
	FileSetTime, tmpDate.YYYY . tmpDate.MM . tmpDate.DD, %holterDir%%filenameOut%.pdf, C	; set the time of PDF in holterDir to 000000 (processed)
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
		} else {
			MsgBox No match
		}
		return
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
	}
	
	if (RegExMatch(fldval["ordering"], "Oi)(Chun)|(Salerno)|(Seslar)", epOrder))  {
		ymatch := epOrder.value()
	}
	
	fileOut1 .= ",""EP_read"",""EP_date"",""MA"""
	fileOut2 .= ",""" ymatch """,""" niceDate(dlDate) """,""" user """"
return
}

Holter:
{
	monType := "H"
	demog := columns(newtxt,"PATIENT\s*DEMOGRAPHICS","Heart Rate Data",1,"Reading Physician")
	holtVals := columns(newtxt,"Medications","INTERPRETATION",,"Total VE Beats")
	
	gosub checkProc												; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	/* Holter PDF is valid. OK to process.
	 * Pulls text between field[n] and field[n+1], place in labels[n] name, with prefix "dem-" etc.
	 */
	fields[1] := ["Last Name", "First Name", "Middle Initial", "ID Number", "Date Of Birth", "Sex"
		, "Source", "Billing Code", "Recorder Format", "Pt\s*?Home\s*?(Phone)?\s*?#?", "Hookup Tech", "Pacemaker\s*?Y/N.", "Medications"
		, "Physician", "Scanned By", "Reading Physician"
		, "Test Date", "Analysis Date", "Hookup Time", "Recording Time", "Analysis Time", "Reason for Test", "Group"]
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
	fileOut2 .= ",""Philips Holter"""
	
return
}

CheckProc:
{
	chk1 := trim(strX(demog,"Last Name",1,9,"First Name",1,10,nn)," `r`n")						; NameL				must be [A-Z]
	chk2 := trim(strX(demog,"First Name",nn,10,"Middle Initial",1,14,nn)," `r`n")				; NameF				must be [A-Z]
	chk3 := trim(strX(demog,"ID Number",nn,9,"Date of Birth",1,13,nn)," `r`n")					; MRN
	chk4 := trim(strX(demog,"Source",nn,7,"Billing Code",1,12,nn)," `r`n")						; Location			must be in SiteVals
	chk5 := trim(strX(demog,"Billing Code",nn,13,"Recorder Format",1,15,nn)," `r`n")			; Billing code		must be valid number
	chk6 := trim(strX(demog,"Physician",nn,10,"Scanned By",1,10,nn)," `r`n")					; Ordering MD
	chk7 := trim(strX(demog,"Test Date",nn,10,"Analysis Date",1,13,nn)," `r`n")					; Study date
	chk8 := trim(strX(demog,"Reason for Test",nn,16,"Group",1,5,nn)," `r`n")					; Indication
	
	if (!(chk1~="[a-z]+")															; Check field values to see if proper demographics
		&& !(chk2~="[a-z]+") 
		&& (chk4~="i)(CRD|EKG|ECO|DCT|Outpatient|Inpatient|Emergency|Day Surg)") 
		&& (chk5~="\d{8}"))
	{
		return																		; All tests valid, return to processing Holter
	}
	else 																			; Not valid PDF, get demographics post hoc
	{
		Clipboard := chk1 ", " chk2													; can just paste into CIS search bar
		MsgBox, 4096,, % "Validation failed for:`n   " chk1 ", " chk2 "`n   " chk3 "`n   " chk4 "`n   " chk5 "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
		ptDem := Object()
		ptDem["nameL"] := chk1														; Placeholder values for fetchGUI from PDF
		ptDem["nameF"] := chk2
		ptDem["mrn"] := chk3
		ptDem["Loc"] := chk4
		ptDem["Account number"] := chk5
		ptDem["Provider"] := trim(RegExReplace(chk6,"i)$Dr\.? "))
		ptDem["EncDate"] := chk7
		ptDem["Indication"] := chk8
		
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
	}
	return
}

Zio:
{
	monType:="Z"
	newTxt:=""
	Loop, parse, maintxt, `n,`r									; first pass, clean up txt
	{
		i:=A_LoopField
		if !(i)													; skip entirely blank lines
			continue
		newTxt .= i . "`n"
	}
	FileDelete tempfile.txt
	FileAppend %newtxt%, tempfile.txt
	
	zdat := columns(newtxt,"","Preliminary Findings",,"Enrollment Period")
	znam := trim(cleanSpace(columns(zdat,"Report for","Date of Birth")))
	fieldColAdd("dem","Name_L",strX(znam, "", 1,1, ",", 1,1))
	fieldColAdd("dem","Name_F",strX(znam, ",", 1,2, "", 1,1))

	zdem := columns(zdat,"Date of birth","Ventricular Tachycardia",1,"Patient ID","Gender","Primary Indication")
	fields[1] := ["Date of Birth","Prescribing Clinician","Patient ID","Managing Location","Gender","Primary Indication"]
	labels[1] := ["DOB","Ordering","MRN","Site","Sex","Indication"]
	fieldvals(zdem,1,"dem")
	
	zarr := columns(zdat,"Ventricular Tachycardia (4 beats or more)","iRhythm Technologies, Inc.",1)
	fields[2] := ["Ventricular Tachycardia (4 beats or more)","Supraventricular Tachycardia (4 beats or more)"
		,"Pauses (3 secs or longer)","Atrial Fibrillation","AV Block (2nd"]
	labels[2] := ["VT","SVT","Pauses","AF","AVBlock"]
	fieldvals(zarr,2,"arr")
	
	znums := columns(zdat,"Enrollment Period","",1)
	
	ztime := columns(znums,"Enrollment Period","Heart Rate",1,"Analysis Time")
	fields[3] := ["Enrollment Period","Analysis Time"]
	labels[3] := ["Enrolled","Analyzed"]
	fieldvals(ztime,3,"time")
	
	zrate := columns(znums,"Heart Rate","Patient Events",1)
	fields[4] := ["Maximum HR","Minimum HR","Average HR"]
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
	
	zsve := columns(znums,"Supraventricular Ectopy (SVE/PACs)","Ventricular Ectopy (VE/PVCs)")
	fields[7] := ["Isolated","Couplet","Triplet"]
	labels[7] := ["Single","Couplets","Triplets"]
	fieldvals(zsve,7,"sve")
	
	zve := columns(znums,"Ventricular Ectopy (VE/PVCs)","")
	fields[8] := ["Isolated","Couplet","Triplet","Longest Ventricular Bigeminy Episode","Longest Ventricular Trigeminy Episode"]
	labels[8] := ["Single","Couplets","Triplets","LongestBigem","LongestTrigem"]
	fieldvals(zve,8,"ve")
	
	zinterp := cleanspace(columns(newtxt,"Preliminary Findings","SIGNATURE",,"Final Interpretation"))
	zinterp := trim(StrX(zinterp,"",1,0,"Final Interpretation",1,20))
	fileout1 .= """INTERP""`n"
	fileout2 .= """" . zinterp . """`n"

return
}

EventRec:
{
	fields := ["PATIENT INFORMATION","Name:","ID #:","4800 SAND POINT","DOB:","Sex:","Phone:"
		,"Monitor Type:","Diag:","Delivery Code:","Enrollment Period:","Date"
		,"SYMPTOMS:","ACTIVITY:","FINDINGS:","COMMENTS:","EVENT RECORDER DATA:"]
	n:=0
	Loop, parse, maintxt, `n,`r
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

columns(x,blk1,blk2,incl:="",col2:="",col3:="",col4:="") {
/*	Returns string as a single column.
	x 		= input string
	blk1	= leading string to start block
	blk2	= ending string to end block
	incl	= if null, exclude blk1 string; if !null, remove blk1 string
	col2	= string demarcates start of COLUMN 2
	col3	= string demarcates start of COLUMN 3
	col4	= string demarcates start of COLUMN 4
*/
	txt := stRegX(x,blk1,1,(incl ? 0 : StrLen(blk1)),blk2,StrLen(blk2))
	StringReplace, col2, col2, %A_space%, [ ]+, All
	StringReplace, col3, col3, %A_space%, [ ]+, All
	StringReplace, col4, col4, %A_space%, [ ]+, All
	
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

fieldvals(x,bl,bl2) {
/*	Matches field values and results. Gets text between FIELDS[k] to FIELDS[k+1]. Excess whitespace removed. Returns results in array BLK[].
	x	= input text
	bl	= which FIELD number to use
	bl2	= label prefix
*/
	global fields, labels, fldval
	
	for k, i in fields[bl]
	{
		j := fields[bl][k+1]
		m := trim(stRegX(x,i,n,1,j,1,n), " `n")
		lbl := labels[bl][A_index]
		cleanSpace(m)
		cleanColon(m)
		fldval[lbl] := m
		formatField(bl2,lbl,m)
	}
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
	BS .= "(.*?)\s{3}"
	rem:="[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"
	pos0 := RegExMatch(h,((BS~=rem)?"Oim"BS:"Oim)"BS),bPat,((BO<1)?1:BO))
	pos1 := RegExMatch(h,((ES~=rem)?"Oim"ES:"Oim)"ES),ePat,pos0+bPat.len)
	N := pos1+((ET)?0:(ePat.len))
	return substr(h,pos0+((BT)?(bPat.len):0),N-pos0-bPat.len)
}

formatField(pre, lab, txt) {
	global monType, Docs
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
	
;	ZIO patch specific search fixes
	if (monType="Z") {
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

parseDate(x) {
; Disassembles "2/9/2015" or "2/9/2015 8:31" into Yr=2015 Mo=02 Da=09 Hr=08 Min=31
	StringSplit, DT, x, %A_Space%
	StringSplit, DY, DT1, /
	StringSplit, DHM, DT2, :
	return {"MM":zDigit(DY1), "DD":zDigit(DY2), "YYYY":DY3, "hr":zDigit(DHM1), "min":zDigit(DHM2), "Date":DT1, "Time":DT2}
}

niceDate(x) {
	if !(x)
		return error
	FormatTime, x, %x%, MM/dd/yyyy
	return x
}

zDigit(x) {
; Add leading zero to a number
	return SubStr("0" . x, -1)
}

#Include CMsgBox.ahk
#Include xml.ahk
#Include sift3.ahk
