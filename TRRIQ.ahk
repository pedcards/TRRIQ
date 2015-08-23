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
IfInString, fileDir, Dropbox					; Change enviroment if run from development vs production directory
{
	isAdmin := true
	holterDir := ".\Holter PDFs\"
	importFld := ".\Import\"
	chipDir := ".\Chipotle\"
} else {
	isAdmin := false
	holterDir := "\\chmc16\Cardio\EP\Holter\Holter PDFs\"
	importFld := "\\chmc16\Cardio\EP\Holter\Import\"
	chipDir := "\\chmc16\Cardio\Inpatient List\Chipotle\"
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
	if ((tmp1="SCH") or (tmp1="FELLOWS")) {
		tmpGrp:=tmp1
		tmpChk:=true
		tmpIdx:=0
		continue
	}
	if !(tmp1) {
		continue
	}
	if !(tmpChk) {
		continue
	}
	tmpIdx += 1
	StringSplit, tmpPrv, tmp1, %A_Space%`"
	tmpPrv := substr(tmpPrv1,1,1) . ". " . tmpPrv2				; F. Last
	tmpPrv := tmpPrv2 ", " tmpPrv1								; Last, First
	Docs[tmpGrp,tmpIdx] := tmpPrv
	Docs[tmpGrp ".eml",tmpIdx] := tmp4
}
;MsgBox,,% Docs[tmpGrp,tmpIdx], % Docs[tmpGrp ".eml",tmpIdx]
;MsgBox % Docs["SCH.eml",ObjHasValue(Docs["SCH"],"Chun, Terry")]
;ExitApp

demVals := ["MRN","Account Number","DOB","Age","Sex","Loc","Provider","PCP"]

if (%0%) {										; For each parameter,
	fileIn := %1%								; Gets parameter dropped/passed to script/exe
	phase := "Process PDF"
}

if !(phase) {
	phase := CMsgBox("Which task?","","*&Enter Holter|&Process PDF","Q","")
}
if (phase = "Enter Holter") {
	gosub fetchDem
	ExitApp
}
if (phase = "Process PDF") {
	if (fileIn) {
		splitpath, fileIn,,,,fileNam
		gosub MainLoop
		ExitApp
	}
	loop, %holterDir%*.pdf
	{
		fileIn := A_LoopFileFullPath
		;MsgBox % fileIn
		gosub MainLoop
	}

	ExitApp

}

ExitApp

FetchDem:
{
	ptDem := Object()
	mdX := Object()
	mdY := Object()
	ptDem["bit"] := 0
	getDem := true
	gosub fetchGUI
	while (getDem) {									; Repeat until we get tired of this
		clipboard :=
		ClipWait
		if !ErrorLevel {
			clk := parseClip(clipboard)
			if !ErrorLevel {
				MouseGetPos, mouseXpos, mouseYpos, mouseWinID, mouseWinClass, 2
				ptDem[clk.field] := (clk.value) ? clk.value : ptDem[clk.field]
				ptDem["bit"] := (clk.bit) ? (ptDem["bit"] |= 2**clk.bit) : ptDem["bit"]
				if (clk.field = "Provider") {
					mdX[4] := mouseXpos
					mdY[1] := mouseYpos
					mdProv := true
				}
				if (clk.field = "Account Number") {
					mdX[1] := mouseXpos
					mdY[3] := mouseYpos
					mdAcct := true
				}
				if (mdProv and mdAcct) {
					mdXd := (mdX[4]-mdX[1])/3
					mdX[2] := mdX[1]+mdXd
					mdX[3] := mdX[2]+mdXd
					mdY[2] := mdY[1]+(mdY[3]-mdY[1])/2
					tmp := mouseGrab((mdX[3]-mdX[1])*0.95,mdY[1])
					ptDem["nameL"] := strX(tmp,,1,0, ",",1,1)
					ptDem["nameF"] := strX(tmp,",",1,2, "",1,1)
					ptDem["MRN"] := mouseGrab(mdX[1],mdY[2])
					ptDem["DOB"] := mouseGrab(mdX[2],mdY[2])
					ptDem["Sex"] := mouseGrab(mdX[3],mdY[1])
					ptDem["Loc"] := mouseGrab(mdX[3]+mdXd*0.5,mdY[2])
					tmp := mouseGrab(mdX[3],mdY[3])
					ptDem["Type"] := strX(tmp,,1,0, " [",1,2)
					ptDem["EncDate"] := strX(tmp," [",1,2, " ",1,1)
					mdProv := false
					mdAcct := false
				}
			}
		}
		gosub fetchGUI							; Update GUI with new info
		
		; This would be a good place to inject the data to the Lifewatch program
	}
	return
}

mouseGrab(x,y) {
	MouseMove, %x%, %y%, 0
	Click 2
	sleep 100
	ClipWait
	clk := parseClip(clipboard)
	return clk.value
	
}

parseClip(clip) {
	global demVals
	StringSplit, val, clip, :
	if (pos:=ObjHasValue(demVals, val1)) {
		return {"field":val1, "value":val2, "bit":pos}
	}
	if (RegExMatch(clip,"O)(Outpatient)|(Inpatient)\s\[",valMatch)) {
		return {"field":"Type", "value":clip}
	}
	if (RegExMatch(clip,"O)[A-Z\-\s]*, [A-Z\-]*",valMatch)) {
		return {"field":"Name", "value":valMatch.value()}
	}
	return Error
}

fetchGUI:
{
	fYd := 30,	fXd := 80
	fX1 := 12,	fX2 := fX1+fXd
	fW1 := 60,	fW2 := 190
	fH := 20
	fY := 10
	fTxt := "To auto-grab demographic info:`n"
		.	"	1) Double-click Encounter Number`n"
		.	"	2) Double-click Provider"
	Gui, fetch:Destroy
	Gui, fetch:Add, Text, % "x" fX1 , % fTxt	
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd*2) " w" fW1 " h" fH , First
	Gui, fetch:Add, Edit, % "x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["nameF"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Last
	Gui, fetch:Add, Edit, % "x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["nameL"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , MRN
	Gui, fetch:Add, Edit, % "x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["MRN"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , DOB
	Gui, fetch:Add, DateTime, % "x" fX2 " y" fY-4 " w" fW2 " h" fH, % ptDem["DOB"], MM/dd/yyyy
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Date placed
	Gui, fetch:Add, DateTime, % "x" fX2 " y" fY-4 " w" fW2 " h" fH, % ptDem["EncDate"], MM/dd/yyyy
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Encounter #
	Gui, fetch:Add, Edit, % "x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["Account Number"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Ordering MD
	Gui, fetch:Add, Edit, % "x" fX2 " y" fY-4 " w" fW2 " h" fH , % ptDem["Provider"]
;	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH , Your initials
;	Gui, fetch:Add, Edit, % "x" fX2 " y" fY-4 " w" fW2 " h" fH , % user
	Gui, fetch:Add, Button, % "x" fX1+10 " y" (fY += fYD) " h" fH+10 " w" fW1+fW2 " gfetchSubmit", Submit!
	Gui, fetch:Show, AutoSize, Enter Demographics
	return
}

fetchGuiClose:
ExitApp

fetchSubmit:
{
/* some error checking
	Check for required elements
	Error check and normalize Ordering MD name
	Check for Lifewatch exe
	Fill Lifewatch data and submit
	The repeat the cycle
*/
	gosub fetchDem
	return
}

MainLoop:
{
	RunWait, pdftotext.exe -l 2 -table -fixed 3 "%fileIn%" temp.txt
	FileRead, maintxt, temp.txt
	blocks := Object()
	fields := Object()
	fldval := {}
	labels := Object()
	newTxt := Object()
	blk := Object()
	blk2 := Object()
	fileOut1 := fileOut2 := ""
	summBl := summ := ""

	Loop, parse, maintxt, `n,`r,%A_Space%					; Identify filetype by text in first lines
	{
		i:=A_LoopField										; Search for first case insensitive string match
		if (InStr(i,"Holter Ext")=1) {
			gosub Holter
			break
		}
		;~ if (InStr(i,"TRANSTELEPHONIC ARRHYTHMIA")=1) {
			;~ gosub EventRec
			;~ break
		;~ }
		if (RegExMatch(i,"i)zio.*xt.*PATCH")) {
			gosub Zio
			break
		}
		if A_Index>9												; only search in the first several lines
			Break													; might be possible to accomplish this with
	}																; i:=substr(maintxt,1,1024)

	gosub epRead
	fileOut1 .= (substr(fileOut1,0,1)="`n") ?: "`n"
	fileOut2 .= (substr(fileOut2,0,1)="`n") ?: "`n"
	fileout := fileOut1 . fileout2
	tmpDate := parseDate(fldval["Test_Date"])
	filenameOut := fldval["Name_L"] " " fldval["MRN"] " " tmpDate.MM tmpDate.DD tmpDate.YYYY
	;MsgBox % filenameOut
	FileDelete, %importFld%%fileNameOut%.csv
	FileAppend, %fileOut%, %importFld%%fileNameOut%.csv
	FileMove, %fileIn%, %holterDir%%filenameOut%.pdf, 1
Return
}

epRead:
{
	FileGetTime, dlDate, %fileIn%
	FormatTime, dlDay, %dlDate%, dddd
	if (dlDay="Friday") {
		dlDate += 3, Days
	}
	FormatTime, dlDate, %dlDate%, yyyyMMdd
	
	y := new XML(chipDir "currlist.xml")
	RegExMatch(y.selectSingleNode("//call[@date='" dlDate "']/EP").text, "Oi)(Chun)|(Salerno)|(Seslar)", ymatch)
	if !(ymatch := ymatch.value()) {
		ymatch := epMon ? epMon : cmsgbox("Electronic Forecast not complete","Which EP on Monday?","Chun|Salerno|Seslar","Q")
		epMon := ymatch
	}
	
	if (RegExMatch(fldval["ordering"], "Oi)(Chun)|(Salerno)|(Seslar)", epOrder))  {
		ymatch := epOrder.value()
	}
	
	fileOut1 .= ",""EP_read"""
	fileOut2 .= ",""" ymatch """"
return
}

Holter:
{
	monType := "H"
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
	
	demog := columns(newtxt,"PATIENT DEMOGRAPHICS","Heart Rate Data",1,"Reading Physician")
	holtVals := columns(newtxt,"Medications","INTERPRETATION",,"Total VE Beats")

	fields[1] := ["Last Name", "First Name", "Middle Initial", "ID Number", "Date Of Birth", "Sex"
		, "Source", "Billing Code", "Recorder Format", "Pt(.*?)Home(.*?)#", "Hookup Tech", "Pacemaker(.*?)Y/N.", "Medications"
		, "Physician", "Scanned By", "Reading Physician"
		, "Test Date", "Analysis Date", "Hookup Time", "Recording Time", "Analysis Time", "Reason for Test", "Group"]
	labels[1] := ["Name_L", "Name_F", "Name_M", "MRN", "DOB", "Sex"
		, "Site", "Billing", "Device_SN", "VOID", "Hookup_tech", "VOID", "Meds"
		, "Ordering", "Scanned_by", "Reading"
		, "Test_date", "Scan_date", "Hookup_time", "Recording_time", "Analysis_time", "Indication", "VOID"]
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
	
	tmp := columns(RegExReplace(newtxt,"i)technician.*comments?:","TECH COMMENT:"),"TECH COMMENT:","Signed :")
	StringReplace, tmp, tmp, .`n , .%A_Space% , All
	fileout1 .= """INTERP"""
	fileout2 .= """" cleanspace(trim(tmp," `n")) """"
	
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
	fields[5] := ["Number of Triggered Events:","Findings within ± 45 sec of Triggers:","Number of Diary Entries:","Findings within ± 45 sec of Entries:"]
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
	txt := strX(x,blk1,1,(incl ? 0 : StrLen(blk1)),blk2,1,StrLen(blk2))
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
	txt:=RegExReplace(txt,"i)BPM|Event(s)?|Beat(s)?|( sec(s)?)|\(.*%\)")	; 	Remove units from numbers
	txt:=RegExReplace(txt,"(:\d{2}?)(AM|PM)","$1 $2")						;	Fix time strings without space before AM|PM
	txt := trim(txt)
	
	if (lab="Ordering") {
		tmpCrd := checkCrd(txt)
		fieldColAdd(pre,lab,tmpCrd.best)
		fieldColAdd(pre,lab "_grp",tmpCrd.group)
		fieldColAdd(pre,lab "_eml",Docs[tmpCrd.Group ".eml",ObjHasValue(Docs[tmpCrd.Group],tmpCrd.best)])
		return
	}
	
;	Lifewatch Holter specific search fixes
	if (monType="H") {
		if InStr(txt," at ") {												;	Split timed results "139 at 8:31:47 AM" into two fields
			tx1 := strX(txt,,1,1," at ",1,4,n)								;		labels e.g. xxx and xxx_time
			tx2 := SubStr(txt,n+4)											;		result e.g. "139" and "8:31:47 AM"
			fieldColAdd(pre,lab,tx1)
			fieldColAdd(pre,lab "_time",tx2)
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

#Include strx.ahk
#Include CMsgBox.ahk
#Include xml.ahk
#Include sift3.ahk