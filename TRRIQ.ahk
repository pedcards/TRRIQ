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
#SingleInstance Force  ; only allow one running instance per user
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

progress,,,TRRIQ intializing...
FileInstall, pdftotext.exe, pdftotext.exe
FileInstall, pdftk.exe, pdftk.exe
FileInstall, libiconv2.dll, libiconv2.dll

SplitPath, A_ScriptDir,,fileDir
user := A_UserName
IfInString, fileDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	chip := httpComm("","full")
	FileDelete, .\Chipotle\currlist.xml
	FileAppend, % chip, .\Chipotle\currlist.xml
	isAdmin := true
	holterDir := ".\Holter PDFs\"
	importFld := ".\Import\"
	chipDir := ".\Chipotle\"
	OnbaseDir1 := ".\Onbase\"
	OnbaseDir2 := ".\HCClinic\"
	webUploadDir := ".\files\Web Upload Files for h3.preventice.com  WebUploadApplication.application\"
	eventlog(">>>>> Started in DEVT mode.")
} else {
	FileGetTime, tmp, TRRIQ.exe
	isAdmin := false
	holterDir := "..\Holter PDFs\"
	importFld := "..\Import\"
	chipDir := "\\childrens\files\HCChipotle\"
	OnbaseDir1 := "\\childrens\apps$\OnbaseFaxFiles\CardiacCathReport\" 
	OnbaseDir2 := "\\childrens\files\HCClinic\Holter Monitors\Holter HIM uploads\"
	webUploadDir := "C:\Web Upload Files for h3.preventice.com  WebUploadApplication.application\"
	eventlog(">>>>> Started in PROD mode. Exe ver " substr(tmp,1,12))
}

/*	Read outdocs.csv for Cardiologist and Fellow names 
*/
progress,,,Scanning providers...
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
	if !(tmp4~="i)(seattlechildrens.org|washington.edu)") {		; skip non-SCH or non-UW providers
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
if fileexist("worklist.xml") {
	wq := new XML("worklist.xml")
} else {
	wq := new XML("<root/>")
	wq.addElement("pending","/root")
	wq.addElement("done","/root")
	wq.save("worklist.xml")
}
scanTempfiles()

demVals := ["MRN","Account Number","DOB","Sex","Loc","Provider"]						; valid field names for parseClip()
sites := "MAIN|BELLEVUE|EVERETT|TRI-CITIES|WENATCHEE|YAKIMA|GREAT FALLS"				; sites we are tracking
sites0 := "TACOMA|SILVERDALE"															; sites we are not tracking
sitesLong := {CRD:"MAIN"
			, EKG:"MAIN"
			, INPATIENT:"MAIN"
			, CRDBEL:"BELLEVUE"
			, CRDEVT:"EVERETT"
			, CRDTRI:"TRI-CITIES"
			, CRDWEN:"WENATCHEE"
			, CRDYAK:"YAKIMA"
			, CRDMT:"GREAT FALLS"
			, CRDTAC:"TACOMA"
			, CRDSIL:"SILVERDALE"}

Loop
{
	Gosub PhaseGUI
	WinWaitClose, TRRIQ Dashboard

	if (phase="Enroll") {
		eventlog("Update Preventice enrollments.")
		gosub CheckPrEnroll
	}

	if (phase="PDF") {
		eventlog("Start PDF folder scan.")
		holterLoops := 0								; Reset counters
		holtersDone := 
		loop, %holterDir%*.pdf							; Process all PDF files in holterDir
		{
			fileNam := RegExReplace(A_LoopFileName,"i)\.pdf")				; fileNam is name only without extension, no path
			fileIn := A_LoopFileFullPath									; fileIn has complete path \\childrens\files\HCCardiologyFiles\EP\HoltER Database\Holter PDFs\steve.pdf
			FileGetTime, fileDt, %fileIn%, C								; fildDt is creatdate/time 
			if (substr(fileDt,-5,2)<4) {									; skip files with creation TIME 0000-0359 (already processed)
				;~ eventlog("Skipping file """ fileNam ".pdf"", already processed.")	; should be more resistant to DST. +0100 or -0100 will still be < 4
				continue
			}
			FileGetSize, fileInSize, %fileIn%
			eventlog("Processing """ fileNam ".pdf"" (" thousandsSep(fileInSize) ").")
			gosub MainLoop													; process the PDF
			if (fetchQuit=true) {											; [x] out of fetchDem means skip this file
				continue
			}
			if !IsObject(ptDem) {											; bad file, never acquires demographics
				continue
			}
			;~ FileDelete, %fileIn%
			holterLoops++													; increment counter for processed counter
			holtersDone .= A_LoopFileName "->" filenameOut ".pdf`n"			; add to report
		}
		MsgBox % "Monitors processed (" holterLoops ")`n" holtersDone
		FileCopy, .\logs\fileWQ.csv, %chipDir%fileWQ-copy.csv, 1
	}
	if (phase="Upload") {
		eventlog("Start Mortara preparation/upload.")
		MortaraUpload()
	}
}

ExitApp

PhaseGUI:
{
	phase :=
	Gui, phase:Destroy
	Gui, phase:Default
	Gui, +AlwaysOnTop

	Gui, Add, Text, x650 y20 w200 h110
		, % "Patients registered in Preventice (" wq.selectNodes("/root/pending/enroll").length ")`n"
		.	"Last Enrollments update: " niceDate(wq.selectSingleNode("/root/pending").getAttribute("update")) 
		;~ .	"Dbl-click a patient item to:`n"
		;~ .	"  - Log upload to Preventice`n"
		;~ .	"  - Note communication`n"
		;~ .	"  - Delete a record"
	Gui, Add, GroupBox, x640 y0 w220 h60
	
	Gui, Font, Bold
	Gui, Add, Button
		, Y+20 w220 h40 vPDF gPhaseTask
		, Process PDF folder
	Gui, Add, Button
		, Y+20 wp h40 vEnroll gPhaseTask
		, Grab Preventice enrollments
	Gui, Add, Button
		, Y+20 wp h40 vUpload gPhaseTask
		, Prepare/Upload Holter
	Gui, Font, Normal
	
	Gui, Add, Tab3, -Wrap x10 y10 w620 h240 vWQtab, % "ALL|" sites						; add Tab bar with tracked sites
	WQlist()
	
	Menu, menuSys, Add, Scan tempfiles, scanTempFiles
	Menu, menuSys, Add, Find returned devices, WQfindlost
	Menu, menuSys, Add, Find close matches, WQfindclose
	Menu, menuSys, Add, Show GeoIP info, showGeoIP
	Menu, menuHelp, Add, About TRRIQ, menuTrriq
	Menu, menuHelp, Add, Instructions..., menuInstr
		
	Menu, menuBar, Add, System, :menuSys
	Menu, menuBar, Add, Help, :menuHelp
	
	Gui, Menu, menuBar
	Gui, Show,, TRRIQ Dashboard
	return
}

PhaseGUIclose:
{
	wq.save("worklist.xml")
	eventlog("<<<<< Session end.")
	ExitApp
}	

menuTrriq:
{
	Gui, phase:hide
	FileGetTime, tmp, TRRIQ.exe
	MsgBox, 64, About..., % "TRRIQ version " substr(tmp,1,12) "`nTerrence Chun, MD"
	Gui, phase:show
	return
}
menuInstr:
{
	Gui, phase:hide
	MsgBox How to...
	gui, phase:show
	return
}

PhaseTask:
{
	phase := A_GuiControl
	Gui, phase:Hide
	return
}

WQtask() {
	agc := A_GuiControl
	if !instr(agc,"WQlv") {
		return
	}
	if !(A_GuiEvent="DoubleClick") {
		return
	}
	Gui, ListView, %agc%
	LV_GetText(idx, A_EventInfo,1)
	if (idx="ID") {
		return
	}
	global wq, user
	
	Gui, phase:Hide
	pt := readWQ(idx)
	idstr := "/root/pending/enroll[@id='" idx "']"
	
	choice := cmsgbox("Patient task"
			,	"Which action on this patient?`n`n"
			.	pt.Name "`n"
			.	"  MRN: " pt.MRN "`n"
			.	"  Date: " niceDate(pt.date) "`n"
			.	"  Provider: " pt.prov "`n"
			.	strQ(pt.FedEx,"  FedEx: ###`n")
			, "NOTE communication|"
			. "Log UPLOAD to Preventice|"
			. "Mark record as DONE"
			, "Q")
	if (choice="Close") {
		return
	}
	if instr(choice,"upload") {
		InputBox ,inDT,Upload log,`n`nEnter date uploaded to Preventice,,,,,,,,% niceDate(A_now)
		if (ErrorLevel) {
			return
		}
		tmp := parseDate(inDT)
		dt := tmp.YYYY tmp.MM tmp.DD
		if !IsObject(wq.selectSingleNode(idstr "/sent")) {
			wq.addElement("sent",idstr)
		}
		wq.setText(idstr "/sent",dt)
		wq.setAtt(idstr "/sent",{user:user})
		wq.save("worklist.xml")
		eventlog(pt.MRN " " pt.Name " study " pt.Date " uploaded to Preventice.")
		MsgBox, 4160, Logged, % pt.Name "`nUpload date logged!"
		return
	}
	if instr(choice,"note") {
		list :=
		Loop, % (notes:=wq.selectNodes(idstr "/notes/note")).length 
		{
			k := notes.item(A_index-1)
			list .= k.getAttribute("date") "/" k.getAttribute("user") ": " k.text "`n"
		}
		note := maxinput("Communication note", list "`nEnter a brief communication note",60)
		if (note="") {
			return
		}
		if !IsObject(wq.selectSingleNode(idstr "/notes")) {
			wq.addElement("notes",idstr)
		}
		if (RegExMatch(note,"((\d\s*){12})",fedex)) {
			MsgBox,4132,, % "FedEx tracking number?`n" fedex1
			IfMsgBox, Yes
			{
				fedex := RegExReplace(fedex1," ")
				if !IsObject(wq.selectSingleNode(idstr "/fedex")) {
					wq.addElement("fedex",idstr)
				}
				wq.setText(idstr "/fedex",fedex)
				wq.setAtt(idstr "/fedex", {user:user, date:substr(A_now,1,8)})
				eventlog(pt.MRN "[" pt.Date "] FedEx tracking #" fedex)
			}
		}
		wq.addElement("note",idstr "/notes",{user:user, date:substr(A_now,1,8)},note)
		;~ wq.save("worklist.xml")
		WriteOut("/root/pending","enroll[@id='" idx "']")
		eventlog(pt.MRN "[" pt.Date "] Note from " user ": " note)
		return
	}
	if instr(choice,"done") {
		reason := cmsgbox("Reason"
				, "What is the reason to remove this record from the active worklist?"
				, "Report in CIS|"
				. "Device missing|"
				. "Other (explain)"
				, "E")
		if (reason="Close") {
			return
		}
		if instr(reason,"Other") {
			reason := maxinput("Clear record from worklist","Enter the reason for moving this record",30)
			if (reason="") {
				return
			}
		}
		if !IsObject(wq.selectSingleNode(idstr "/notes")) {
			wq.addElement("notes",idstr)
		}
		wq.addElement("note",idstr "/notes",{user:user, date:substr(A_now,1,8)},"MOVED: " reason)
		moveWQ(idx)
		eventlog(idx " Move from WQ: " reason)
	}
return	
}

maxinput(title, prompt, max) {
	Loop
	{
		prompt .= "`n(Max " max " chars)"
		StrReplace(prompt,"`n","`n",lines)
		InputBox, reason, % title, % prompt " " lines " lines",,400,% (lines*20)+150
		StringLen, addLength, reason
		If (addLength > max) {
			MsgBox, 0, ERROR, % "String too long. Please explain in less than " max " chars." 	
		} else {
			break
		}
	}
	if (reason="") {
		return error
	}
	
	return reason
}

WQlist() {
	global
	local k, ens, id, e0, now, dt, site
	
	Progress,,,Scanning worklist...
	
	Gui, Add, Listview, -Multi Grid BackgroundSilver W600 H200 gWQtask vWQlv0 hwndHLV0, ID|Enrolled|FedEx|Uploaded|MRN|Enrolled Name|Device|Provider|Site
	
	fileCheck()
	FileOpen(".lock", "W")																; Create lock file.
	loop, parse, sites0, |
	{
		site := A_LoopField
		Loop, % (ens:=wq.selectNodes("/root/pending/enroll[site='" site "']")).length
		{
			k := ens.item(A_Index-1)
			clone := k.cloneNode(true)
			wq.selectSingleNode("/root/done").appendChild(clone)						; copy k.clone to DONE
			k.parentNode.removeChild(k)													; remove k node
			eventlog("Moved " site " record " k.selectSingleNode("mrn").text " " k.selectSingleNode("name").text)
		}
	}
	wq.save("worklist.xml")
	FileDelete, .lock
	
	Loop, parse, sites, |
	{
		i := A_index
		site := A_LoopField
		Gui, Tab, % site
		Gui, Add, Listview, -Multi Grid BackgroundSilver W600 H200 gWQtask vWQlv%i% hwndHLV%i%, ID|Enrolled|FedEx|Uploaded|MRN|Enrolled Name|Device|Provider
		Loop, % (ens:=wq.selectNodes("/root/pending/enroll[site='" site "']")).length
		{
			k := ens.item(A_Index-1)
			id	:= k.getAttribute("id")
			e0 := readWQ(id)
			now := A_Now
			dt := e0.date
			dt -= now, Days
			e0.dev := RegExReplace(e0.dev,"BodyGuardian","BG")
			if (instr(e0.dev,"BG") && (dt > -30)) {
				continue
			}
			Gui, ListView, WQlv%i%
			LV_Add(""
				,id
				,e0.date																;~ ,parseDate(e0.date).MM "/" parseDate(e0.date).DD
				,strQ(e0.fedex,"X")
				,e0.sent																;~ ,strQ(e0.sent,parseDate(e0.date).MM "/" parseDate(e0.date).DD)
				,e0.mrn
				,e0.name
				,e0.dev
				,e0.prov
				,e0.site)
			Gui, ListView, WQlv0
			LV_Add(""
				,id
				,e0.date																;~ ,parseDate(e0.date).MM "/" parseDate(e0.date).DD
				,strQ(e0.fedex,"X")
				,e0.sent																;~ ,strQ(e0.sent,parseDate(e0.date).MM "/" parseDate(e0.date).DD)
				,e0.mrn
				,e0.name
				,e0.dev
				,e0.prov
				,e0.site)
			
		}
		Gui, ListView, WQlv%i%
		LV_ModifyCol()
		LV_ModifyCol(1,"0")
		LV_ModifyCol(2,"60 Desc")
		LV_ModifyCol(2,"Sort")
		LV_ModifyCol(3,"40")
		LV_ModifyCol(4,"60")
		LV_ModifyCol(6,140)
		LV_ModifyCol(8,130)
	}
	Gui, ListView, WQlv0
	LV_ModifyCol()
	LV_ModifyCol(1,"0")
	LV_ModifyCol(2,"60 Desc")
	LV_ModifyCol(2,"Sort")
	LV_ModifyCol(3,"40")
	LV_ModifyCol(4,"60")
	LV_ModifyCol(6,140)
	LV_ModifyCol(8,130)
	progress, off
	return
}

WQfindlost() {
	MsgBox, 4132, Find devices, Scan database for duplicate devices?`n`n(this can take a while)
	IfMsgBox, Yes
	{
		progress,,, Scanning lost devices
		loop
		{
			res := WQfindreturned()
			if (res="clean") {
				eventlog("Device logs clean.")
				break
			}
			moveWQ(res)
		}
	} 
	reload
}

WQfindreturned() {
	global wq
	
	loop, % (ens:=wq.selectNodes("/root/pending/enroll")).length
	{
		e0 := []
		k := ens.item(A_Index-1)
		e0.id := k.getAttribute("id")
		e0.date := k.selectSingleNode("date").text
		enlist .= e0.date "," e0.id "`n"
	}
	sort, enlist
	loop, parse, enlist, `n,`r
	{
		StringSplit, en, A_LoopField, `,
		k := wq.selectSingleNode("/root/pending/enroll[@id='" en2 "']")
		dev := k.selectSingleNode("dev").text
		find := trim(stregX(dev," -",1,1,"$",0))
		findID := en2
		if (find="") {
			continue
		}
		progress,% A_index,, Scanning for reused devices
		loop, parse, enlist, `n,`r
		{
			StringSplit, idk, A_LoopField, `,
			if (idk2=en2) {
				continue
			}
			k2 := wq.selectSingleNode("/root/pending/enroll[@id='" idk2 "']")
			dev2 := k2.selectSingleNode("dev").text
			if instr(dev2,find) {
				found := true
				break
			}
		}
		if (found) {
			return findID
		} 
	}
	return "clean"
}

WQfindclose() {
/*	This may be a redundant function
	Consider removing
*/
	global wq
	
	loop, % (pend := wq.selectNodes("/root/pending/enroll")).Length
	{
		k1 := pend.item(A_Index-1)
		id1 := k1.getAttribute("id")
		e1 := readWQ(id1)
		Progress, % 100*A_index/pend.length,, % e1.date
		
		loop, % (done := wq.selectNodes("/root/done/enroll[date='" e1.date "']")).length				; all items matching [date]
		{
			k2 := done.item(A_index-1)
			id2 := k2.getAttribute("id")
			e2 := readWQ(id2)
			e2.fuzzName := 100*(1-fuzzysearch(e2.name,e1.name))						; percent match
			e2.fuzzMRN	:= 100*(1-fuzzysearch(e2.mrn,e1.mrn))
			if ((e2.fuzzName>85)||(e2.fuzzMRN>85)) {									; close match for either NAME or MRN
				e2.match := id2
				break
			}
		}
		if (e2.match) {
			;~ eventlog("Enrollment close match (" res.mrn "/" e0.mrn ") and (" res.name "/" e0.name ") found in " e0.match "[" date "].")
			MsgBox % "Close match (" e1.mrn "/" e2.mrn ") and (" e1.name "/" e2.name ") found in " e2.match
			e2.match := ""
			continue
		}
	}
	progress, off
	return
}

showGeoIP() {
	geo := httpComm("http://api.geoiplookup.net")
	MsgBox % "IP: " A_IPAddress1 "`n"
		.	"-------------`n"
		.	geo
	return
}

readWQ(idx) {
	global wq
	
	res := []
	k := wq.selectSingleNode("//enroll[@id='" idx "']")
	Loop, % (ch:=k.selectNodes("*")).Length
	{
		i := ch.item(A_index-1)
		node := i.nodeName
		val := i.text
		res[node]:=val
	}
	return res
}

FetchDem:
{
	mdCoord := Object()											; clear Mouse Demographics X,Y coordinate arrays
	getDem := true
	mdProv := false
	mdAcct := false
	
	while (getDem) {									; Repeat until we get tired of this
		clk := Object()
		clipboard :=
		ClipWait, 0
		if !ErrorLevel {								; clipboard has data
			MouseGetPos, mouseXpos, mouseYpos, mouseWinID, mouseWinClass, 2				; put mouse coords into mouseXpos and mouseYpos, and associated winID
			clk := parseClip()
			if !ErrorLevel {															; parseClip {field:value} matches valid data
				WinGetActiveStats, mdTitle, mdWinW, mdWinH, mdWinX, mdWinY				; get window coords as well
				mdXd := mdWinW/6														; determine delta X between columns
				
				if (clk.field = "Provider") {
					if (clk.value~="[[:alpha:]]+.*,.*[[:alpha:]]+") {					; extract provider.value to LAST,FIRST (strip MD, PHD, MI, etc)
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
					mdCoord.x4 := mouseXpos													; demographics grid[4,1]
					mdCoord.y1 := mouseYpos
					mdProv := true														; we have got Provider
					gosub getDemName													; extract patient name, MRN from window title 
				}																		; (this is why it must be sister or parent VM).
				if (clk.field = "Account Number") {
					ptDem["Account Number"] := clk.value
					eventlog("MouseGrab Account Number " clk.value ".")
					mdCoord.x1 := mouseXpos													; demographics grid[1,3]
					mdCoord.y3 := mouseYpos
					mdAcct := true														; we have got Acct Number
					gosub getDemName													; extract patient name, MRN
				}
				if (mdProv and mdAcct) {												; we have both critical coordinates
					mdX0 := 50
					mdCoord.x1 := mdX0
					mdCoord.x2 := mdX0 + mdXd
					mdCoord.x3 := mdX0 + mdXd*2
					mdCoord.x4 := mdX0 + mdXd*3
					mdCoord.y2 := mdCoord.y1+(mdCoord.y3-mdCoord.y1)/2									; determine remaning row coordinate
					
					Gui, fetch:hide														; grab remaining demographic values
					BlockInput, On														; Prevent extraneous input
					ptDem["MRN"] := mouseGrab(mdCoord.x1,mdCoord.y2).value
					ptDem["DOB"] := mouseGrab(mdCoord.x2,mdCoord.y2).value
					ptDem["Sex"] := mouseGrab(mdCoord.x3,mdCoord.y1).value
					eventlog("MouseGrab other fields. MRN=" ptDem["MRN"] " DOB=" ptDem["DOB"] " Sex=" ptDem["Sex"] ".")
					
					tmp := mouseGrab(mdCoord.x3,mdCoord.y3)										; grab Encounter Type field
					ptDem["Type"] := tmp.value
					if (ptDem["Type"]="Outpatient") {
						ptDem["Loc"] := mouseGrab(mdCoord.x4-mdX0-30,mdCoord.y2).value				; most outpatient locations are short strings, click the right half of cell to grab location name
					} else {
						ptDem["Loc"] := tmp.loc
					}
					if !(ptDem["EncDate"]) {											; EncDate will be empty if new upload or null in PDF
						ptDem["EncDate"] := tmp.date
					}
					ptDem["Hookup time"] := tmp.time
					
					mdProv := false														; processed demographic fields,
					mdAcct := false														; so reset check bits
					mdCoord := Object()
					
					BlockInput, Off														; Permit input again
					Gui, fetch:show
					eventlog("MouseGrab other fields."
						. " Type=" ptDem["Type"] " Loc=" ptDem["Loc"]
						. " EncDate=" ptDem["EncDate"] " EncTime=" ptDem["Hookup time"] ".")
				}
				mouseXpos := ""
				mouseYpos := ""
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
	MouseMove, %x%, %y%, 0																; Goto coordinates
	Click 2																				; Double-click
	ClipWait, 0																			; sometimes there is delay for clipboard to populate
	clk := parseClip()																	; get available values out of clipboard
	return clk																			; Redundant? since this is what parseClip() returns
}

parseClip() {
/*	If clip matches "val1:val2" format, and val1 in demVals[], return field:val
	If clip contains proper Encounter Type ("Outpatient", "Inpatient", "Observation", etc), return Type, Date, Time
*/
	global demVals
	
	;~ sleep 100
	clip := clipboard
	
	StringSplit, val, clip, :															; break field into val1:val2
	if (ObjHasValue(demVals, val1)) {													; field name in demVals, e.g. "MRN","Account Number","DOB","Sex","Loc","Provider"
		clipboard := ""
		return {"field":val1
				, "value":val2}
	}
	
	dt := strX(clip," [",1,2, "]",1,1)													; get date
	if (clip~="Outpatient\s\[") {														; Outpatient type
		clipboard := ""
		return {"field":"Type"
				, "value":"Outpatient"
				, "loc":"Outpatient"
				, "date":parseDate(dt).date
				, "time":parseDate(dt).time}
	}
	if (clip~="Inpatient|Observation\s\[") {											; Inpatient types
		clipboard := ""
		return {"field":"Type"
				, "value":"Inpatient"
				, "loc":"Inpatient"
				, "date":""}															; can span many days, return blank
	}
	if (clip~="Day Surg.*\s\[") {														; Day Surg type
		clipboard := ""
		return {"field":"Type"
				, "value":"Day Surg"
				, "loc":"SurgCntr"
				, "date":parseDate(dt).date}
	}
	if (clip~="Emergency") {															; Emergency type
		clipboard := ""
		return {"field":"Type"
				, "value":"Emergency"
				, "loc":"Emergency"
				, "date":parseDate(dt).date}
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
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("Account Number","\d{8,}",1), Encounter #
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
{
	Gui, fetch:destroy
	getDem := false																	; break out of fetchDem loop
	fetchQuit := true
	eventlog("Manual [x] out of fetchDem.")
Return
}

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
		eventlog(ptDem.Provider " matches " matchProv.Best " (" (1-matchProv.fuzz)*100 ").")
		ptDem.Provider := matchProv.Best
	}
	ptDem["Account Number"] := EncNum											; make sure array has submitted EncNum value
	FormatTime, EncDt, %EncDt%, MM/dd/yyyy										; and the properly formatted date 06/15/2016
	ptDem.EncDate := EncDt
	ptDemChk := (ptDem["nameF"]~="i)[A-Z\-]+") && (ptDem["nameL"]~="i)[A-Z\-]+") 					; valid names
			&& (ptDem["mrn"]~="\d{6,7}") && (ptDem["Account Number"]~="\d{8,}") 						; valid MRN and Acct numbers
			&& (ptDem["DOB"]~="[0-9]{1,2}/[0-9]{1,2}/[1-2][0-9]{3}") && (ptDem["Sex"]~="^[MF]") 		; valid DOB and Sex
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
			;~ . ((ptDem["nameF"]) ? "" : "First name`n")
			;~ . ((ptDem["nameL"]) ? "" : "Last name`n")
			;~ . ((ptDem["mrn"]) ? "" : "MRN`n")
			;~ . ((ptDem["Account number"]) ? "" : "Account number`n")
			;~ . ((ptDem["DOB"]) ? "" : "DOB`n")
			;~ . ((ptDem["Sex"]) ? "" : "Sex`n")
			;~ . ((ptDem["Loc"]) ? "" : "Location`n")
			;~ . ((ptDem["Type"]) ? "" : "Visit type`n")
			;~ . ((ptDem["EncDate"]) ? "" : "Date Holter placed`n")
			;~ . ((ptDem["Provider"]) ? "" : "Provider`n")
			;~ . "`nREQUIRED!"
			. "First name " ptDem["nameF"] "`n"
			. "Last name " ptDem["nameL"] "`n"
			. "MRN " ptDem["mrn"] "`n"
			. "Account number " ptDem["Account number"] "`n"
			. "DOB " ptDem["DOB"] "`n"
			. "Sex " ptDem["Sex"] "`n"
			. "Location " ptDem["Loc"] "`n"
			. "Visit type " ptDem["Type"] "`n"
			. "Date Holter placed " ptDem["EncDate"] "`n"
			. "Provider " ptDem["Provider"] "`n"
			. "`nREQUIRED!"
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
	eventlog("Indications entered.")
	return
}

getDem:
{
	ptDem := Object()																	; New enroll needs demographics
	gosub fetchGUI																		; Grab it first
	gosub fetchDem
	if (fetchQuit=true) {
		return
	}
	Loop
	{
		if (ptDem.Indication) {															; loop until we have filled indChoices
			break
		}
		gosub indGUI
		WinWaitClose, Enter indications
	}
	
	return
}

CheckPrEnroll:
{
	while !(WinExist("Patient Enrollment"))
	{
		MsgBox,4161,Update Preventice enrollments
			, % "Navigate on Preventice website to:`n`nEnrollment / Submitted Patients`n`n"
			.	"Click OK when ready to proceed"
		IfMsgBox, Cancel
		{
			return
		}
	}
	loop																				; Repeat until determine done
	{
		clip := grabWebpage("Patient Enrollment")										; Page exists, ask to grab
		if !(clip) {
			break																		; Clicked "Cancel", exit out
		}
		if (clip = clip0) {																; Check if this is the same as the last page
			MsgBox,4144,, % "Done already!`n`nClick on 'Next Page'`nbefore proceding."
			IfMsgBox, OK
			{
				continue
			} else {
				break
			}
		}
		if (instr(clip,"Enrollment Queue (Submitted)")) {
			list := clip
			done:=parseEnrollment(list)
			if !(done) {
				MsgBox,4144,, Reached the end of novel records.`n`nYou may exit scan mode.
			}
			clip0 := clip
		} else {
			MsgBox,4112,, Wrong page!`nNavigate to:`n`nEnrollment / Submitted Patients
		}
	}
	return
}

grabWebpage(title) {
/*	Copy text of an open webpage
 *	title = string in window title
 */
	WinActivate, %title%																; activate the browser window when title matches
	MsgBox, 4145, "%title%" grab, Ready to grab!`n`n`[OK] to grab this page`n[CANCEL] to exit
	IfMsgBox, OK
	{
		WinActivate, %title%															; activate the browser window when title matches
		MouseGetPos,mouseX,mouseY														; get mouse coords
			MouseClick, Left, 0, mouseY													; Click off to far side to clear selection
			Send, ^a^c																	; Select All, Copy
			sleep 200																	; need to pause to fill clipboard
			clip := Clipboard
			MouseClick, Left, 0, mouseY												; Click off to far side to clear selection
		MouseMove, mouseX, mouseY														; move back to original coords
		return clip
	} 
	return error
}

parseEnrollment(x) {
	global wq
	
	fileCheck()
	FileOpen(".lock", "W")															; Create lock file.
	Loop
	{
		blk := stregX(x,"Patient Enrollment",n,1,"Dr\..*?[\r\n]",0,n)
		if !(blk) {
			break
		}
		blk := trim(RegExReplace(blk,"[\r\n]+")," `r`n")
		fields := ["^"
				,"\d{6,7}"
				,"\d{1,2}/\d{1,2}/\d{2,4}"
				,"\w"
				,"Dr. "
				,"$"]
		labels := ["name"
				,"mrn"
				,"date"
				,"dev"
				,"prov"
				,"end"]
		res:=scanX(blk,fields,labels)
		tmp := parseDate(res.date)
		date := tmp.YYYY tmp.MM tmp.DD
		count ++
		
		if IsObject(wq.selectSingleNode("/pending/enroll[mrn='" res.mrn "'][dev='" res.dev "']")) {			; S/N is currently in use
			eventlog("Enrollment for " res.dev " already exists in Pending.")
			continue
		}
		if IsObject(ens := wq.selectSingleNode("//enroll[date='" date "'][mrn='" res.mrn "']")) {			; exists in PENDING or DONE
			eventlog("Enrollment for " res.mrn " " res.name " " date " already exists in " ens.parentNode.nodeName ".")
			continue
		} 
		
		loop, % (ens := wq.selectNodes("//enroll[date='" date "']")).length				; all items matching [date]
		{
			k := ens.item(A_index-1)
			e0 := []
			e0.id := k.getAttribute("id")
			e0.name	:= k.selectSingleNode("name").text
			e0.mrn	:= k.selectSingleNode("mrn").text
			e0.fuzzName := 100*(1-fuzzysearch(e0.name,res.name))						; percent match
			e0.fuzzMRN	:= 100*(1-fuzzysearch(e0.mrn,res.mrn))
			if ((e0.fuzzName>85)||(e0.fuzzMRN>85)) {									; close match for either NAME or MRN
				e0.match := k.parentNode.nodeName 
				break
			}
		}
		if (e0.match) {
			eventlog("Enrollment close match (" res.mrn "/" e0.mrn ") and (" res.name "/" e0.name ") found in " e0.match "[" date "].")
			e0.match := ""
			continue
		}
		
		/*	No perfect or close match
		 *	add new record to PENDING
		 */
		sleep 1																			; delay 1ms to ensure different tick time
		id := A_TickCount 
		wq.addElement("enroll","/root/pending",{id:id})
		newID := "/root/pending/enroll[@id='" id "']"
		wq.addElement("date",newID,date)
		wq.addElement("name",newID,res.name)
		wq.addElement("mrn",newID,res.mrn)
		wq.addElement("dev",newID,res.dev)
		wq.addElement("prov",newID,filterProv(res.prov).name)
		wq.addElement("site",newID,filterProv(res.prov).site)
		wq.addElement("webgrab",newID,A_now)
		done ++
		
		eventlog("Added new registration " res.mrn " " res.name " " date ".")
	}
	wq.selectSingleNode("/root/pending").setAttribute("update",A_now)					; set pending[@update] attr
	wq.save("worklist.xml")
	filedelete, .lock
	
	return done
/*		value = records added
 *		null  = no records added (no unique)
 */
}

scanX(txt,fields,labels) {
	res := Object()
	for k, i in fields																	; Step through each val "i" from fields[bl,k]
	{
		x := fields[k]
		y := fields[k+1]
		
		val := stregX(txt,x,n,0,y,1,n)
		
		res[labels[k]]:=trim(val)
	}
	return res
}

findWQid(DT,MRN,name="") {
	global wq
	
	if IsObject(x := wq.selectSingleNode("//enroll[date='" DT "'][mrn='" MRN "']")) {				; Perfect match
	} else if IsObject(x := wq.selectSingleNode("//enroll[mrn='" MRN "']")) {						; or matches MRN only
	} else if IsObject(x := wq.selectSingleNode("//enroll[name='" name "']")) {						; or neither, find matching name
	}
	return {id:x.getAttribute("id"),node:x.parentNode.nodeName}										; will return null (error) if no match
}

scanTempfiles() {
	global wq
	count := 0
	
	filecheck()
	FileOpen(".lock", "W")															; Create lock file.
	
	loop, files, tempfiles/*.csv
	{
		filenm := A_LoopFileName
		files ++
		RegExMatch(filenm,"O)^(\d{6,7}) (.*)? (\d{2}-\d{2}-\d{4})",wqnm)
		if !(wqnm.value(0)) {
			continue
		}
		mrn :=  wqnm.value(1)
		name := wqnm.value(2)
		dt := parseDate(wqnm.value(3))
		date := dt.YYYY dt.MM dt.DD
		
		if IsObject(wq.selectSingleNode("/root/done/enroll[mrn='" mrn "'][date='" date "']")) {
			continue
		}
		
		id := A_TickCount 
		wq.addElement("enroll","/root/done",{id:id})
		newID := "/root/done/enroll[@id='" id "']"
		wq.addElement("date",newID,date)
		wq.addElement("name",newID,name)
		wq.addElement("mrn",newID,mrn)
		count ++
		sleep 1
	}
	wq.save("worklist.xml")
	FileDelete, .lock
	eventlog("Scanned " files " files, " count " DONE records added.")
return "Scanned " files " files, " count " DONE records added."
}

MortaraUpload()
{
	global wq, mu_UI, ptDem, fetchQuit, MtCt, webUploadDir, user
	ptDem := Object()
	mu_UI := Object()
	fetchQuit := false
	MtCt := ""
	
	Loop																				; Do until Web Upload program is running
	{
		if (muWinID := winexist("Mortara Web Upload")) {								; Break out of loop when window present
			WinGetClass, muWinClass, ahk_id %muWinID%									; Grab WinClass string for processing
			break
		}
		MsgBox, 262193, Inject demographics, Must launch Mortara Web Upload program!	; Otherwise remind to launch program
		IfMsgBox Cancel 
		{
			return																		; Can cancel out of this process if desired
		}
	}
	
	DetectHiddenText, Off																; Only check the visible text
	Loop																				; Do until either Upload or Prepare window text is present
	{
		WinGetText, muWinTxt, ahk_id %muWinID%											; Should only check visible window
		if instr(muWinTxt,"Recorder S/N") {
			break
		}
		MsgBox, 262193, Inject demographics
			, Select the device activity,`nTransfer or Prepare Holter
		IfMsgBox Cancel 
		{
			return																		; Can cancel out of this process if desired
		}
	}
	DetectHiddenText, On
	ControlGet , Tabnum, Tab, 															; Get selected tab num
		, WindowsForms10.SysTabControl32.app.0.33c0d9d1
		, ahk_id %muWinID%
	SerNum := substr(stregX(muWintxt,"Status.*?[\r\n]+",1,1,"Recorder S/N",1),-6)		; Get S/N on visible page
	SerNum := SerNum ? trim(SerNum," `r`n") : ""
	eventlog("Device S/N " sernum " attached.")
	
	if (Tabnum=1) {																		; TRANSFER RECORDING TAB
		eventlog("Transfer recording selected.")
		mu_UI := MorUIgrab()
		
; 	******************************
		wuDirDate := ""
		wuDirFull := ""
		Loop, files, % WebUploadDir "Data\*", D											; Get the most recently created Data\xxx folder
		{
			loopDate := A_LoopFileTimeModified
			loopName := A_LoopFileLongPath
			if (loopDate>=wuDirDate) {
				wuDirDate := loopDate
				wuDirFull := loopName
			}
		}
		wuDirShort := strX(wuDirFull,"\",0,1,"",0)
		wuDirDT := RegExReplace(strX(wuDirShort,"_",0,1,"",0),"_")
		eventlog("Found WebUploadDir " wuDirShort " [" wuDirDT "]")
		FileReadLine, wuRecord, % wuDirFull "\RECORD.LOG", 1
		wuDirMRN := trim(RegExReplace(wuRecord,"i)Patient ID:"))
		eventlog("Manifest serial number " serNum ", MRN " wuDirMRN)
; 	******************************
		
		wqStr := "/root/pending/enroll[dev='Mortara H3+ - " SerNum "'][mrn='" wuDirMRN "']"
		wqTR:=wq.selectSingleNode(wqStr)
		if IsObject(wqTR.selectSingleNode("acct")) {									; S/N exists, and valid
			pt := readwq(wqTR.getAttribute("id"))
			ptDem["mrn"] := pt.mrn														; fill ptDem[] with values
			ptDem["loc"] := pt.site
			ptDem["date"] := pt.date
			ptDem["Account Number"] := RegExMatch(pt.acct,"([[:alpha:]]+)(\d{8,})",z) ? z2 : pt.acct
			ptDem["nameL"] := strX(pt.name,"",0,1,",",1,1)
			ptDem["nameF"] := strX(pt.name,",",1,1,"",0)
			ptDem["Sex"] := pt.sex
			ptDem["dob"] := pt.dob
			ptDem["Provider"] := pt.prov
			ptDem["Indication"] := pt.ind
			ptDem["loc"] := z1
			ptDem["wqid"] := wqTR.getAttribute("id")
			eventlog("Found valid registration for " pt.name " " pt.mrn " " pt.date)
			MsgBox,, 
			MsgBox, 262193
				, Match!
				, % "Found valid registration for:`n" pt.name "`n" pt.mrn "`n" pt.date
			IfMsgBox, Cancel
			{
				return
			}
		} else {																		; no valid S/N exists
			gosub getDem																; fill ptDem[] with values
			if (fetchQuit=true) {
				fetchQuit:=false
				return
			}
			ptDem["muphase"] := "upload"
			muWqSave(SerNum)
			wqStr := "/root/pending/enroll[dev='Mortara H3+ - " SerNum "'][mrn='" ptDem["mrn"] "']"
		}
		MorUIfill(mu_UI.TRct,muWinID)
		
		Gui, muTm:Add, Progress, w150 h6 -smooth hwndMtCt 0x8
		Gui, muTm:+ToolWindow
		Gui, muTm:Show, AutoSize, Close to cancel upload...
		SetTimer, muTimer, 50
		
		loop
		{
			if FileExist(wuDirFull "\Uploaded.txt") {
				;~ FileRead, wuDirUpload, % wuDirFull "\Uploaded.txt"
				Gui, muTm:Destroy
				settimer, muTimer, off
				break
			}
			if (ptDem.timer) {
				Gui, muTm:Destroy
				eventlog("muTimer closed.")
				settimer, muTimer, off
				return
			}
		}
		
		if !IsObject(wq.selectSingleNode(wqStr "/sent")) {
			wq.addElement("sent",wqStr)
		}
		wq.setText(wqStr "/sent",substr(A_now,1,8))
		wq.setAtt(wqStr "/sent",{user:user})
		WriteOut("/root/pending","enroll[dev='Mortara H3+ - " SerNum "'][mrn='" ptDem["mrn"] "']")
		eventlog(ptDem.MRN " " ptDem.Name " study " ptDem.Date " uploaded to Preventice.")
	}
	
	if (Tabnum=2) {																		; PREPARE MEDIA TAB
		eventlog("Prepare media selected.")
		mu_UI := MorUIgrab()
		
		gosub getDem
		if (fetchQuit=true) {
			fetchQuit:=false
			return
		}
		
		WinActivate, ahk_id %muWinID%
		sleep 500
		ControlGet, clkbut, HWND,, Set Clock...
		sleep 200
		ControlClick,, ahk_id %clkbut%,,,,NA
		WinWaitClose, Set Recorder Time
		
		MorUIfill(mu_UI.PRct,muWinID)
		
		loop
		{
			winget, x, ProcessName, A													; Dialog has no title
			if !instr(x,"WebUpload") {													; so find the WebUpload
				continue
			}
			WinGetText, x, A
			if (x="OK`r`n") {															; dialog that has only "OK`r`n" as the text
				WinGet, finOK, ID, A
				break
			}
		}
		Winwaitclose, ahk_id %finOK%													; Now we can wait until it is closed
		
		ptDem["muphase"] := "prepare"
		muWqSave(SerNum)
	}
	
	return
}

muTmGuiClose:
{
	ptDem.timer := true
	return
}

muTimer:
{
	GuiControl,,% MtCt
	return
}

muWqSave(sernum) {
	global wq, ptDem, user, sitesLong
	
	filecheck()
	FileOpen(".lock", "W")																; Create lock file.
	wqStr := "/root/pending/enroll[dev='Mortara H3+ - " sernum "']"
	loop, % (ens:=wq.selectNodes(wqStr)).length											; Clear all prior instances of this sernum
	{
		i := ens.item(A_index-1)
		enID := i.getAttribute("id")
		enName := i.selectSingleNode("name").text
		enMRN := i.selectSingleNode("mrn").text
		enDate := i.selectSingleNode("date").text
		enSent := i.selectSingleNode("sent").text
		if (enSent) {																	; pending/enroll/sent = uploaded, waiting for PDF
			continue																	; so don't remove
		}
		enStr := "/root/pending/enroll[@id='" enId "']"
			wq.addElement("removed",enStr,{user:A_UserName},A_Now)						; set as done
			x := wq.selectSingleNode(enStr)												; reload x node
			clone := x.cloneNode(true)
			wq.selectSingleNode("/root/done").appendChild(clone)						; copy x.clone to z.DONE
			x.parentNode.removeChild(x)													; remove enStr node
			
			wq.save("worklist.xml")
		eventlog("Device " sernum " reg to " enName " - " enMRN " on " enDate ", moved to DONE list.")
	}
	filedelete, .lock
	
	if (ptDem.EncDate) {
		tmp := parsedate(ptDem.EncDate)
		ptDem.date := tmp.YYYY tmp.MM tmp.DD
	}
	
	id := A_TickCount 
	wq.addElement("enroll","/root/pending",{id:id})
	newID := "/root/pending/enroll[@id='" id "']"
	wq.addElement("date",newID,(ptDem["date"]) ? ptDem["date"] : substr(A_now,1,8))
	wq.addElement("name",newID,ptDem["nameL"] ", " ptDem["nameF"])
	wq.addElement("mrn",newID,ptDem["mrn"])
	wq.addElement("sex",newID,ptDem["Sex"])
	wq.addElement("dob",newID,ptDem["dob"])
	wq.addElement("dev",newID,"Mortara H3+ - " sernum)
	wq.addElement("prov",newID,ptDem["Provider"])
	wq.addElement("site",newID,sitesLong[ptDem["loc"]])										; need to transform site abbrevs
	wq.addElement("acct",newID,ptDem["loc"] ptDem["Account Number"])
	wq.addElement("ind",newID,ptDem["Indication"])
	wq.addElement(ptDem["muphase"],newID,A_now)
	
	writeOut("/root/pending","enroll[@id='" id "']")
	eventlog(ptDem["muphase"] ": " sernum " registered to " ptDem["mrn"] " " ptDem["nameL"] ".") 
	
	return
}

MorUIgrab() {
	id := WinExist("Mortara Web Upload")
	q := Object()
	WinGet, WinText, ControlList, ahk_id %id%

	Loop, parse, % WinText, `n,`r
	{
		str := A_LoopField
		if !(str) {
			continue
		}
		ControlGetText, val, %str%, ahk_id %id%
		ControlGetPos, mx, my, mw, mh, %str%, ahk_id %id%
		if (val=" Transfer Recording ") {
			TRct := A_index
		}
		if (val=" Prepare Recorder Media ") {
			PRct := A_Index
		}
		el := {x:mx,y:my,w:mw,h:mh,str:str,val:val}
		q[A_index] := el
	}
	q.txt := WinText
	q.TRct := TRct
	q.PRct := PRct
	
	return q
}

MorUIfind(val,start) {
/*	val = string to find, e.g. "First Name"
	start = starting index for TR vs PR
	returns object element matching val string
*/
	global mu_UI
	
	loop, % mu_UI.maxIndex()
	{
		if (A_index<start) {
			continue
		}
		el := mu_UI[A_Index]
		if (val=trim(el.val," :")) {
			idx := A_Index
			break
		}
	}
	if !(idx) {
		return
	}
	
	return el
}

MorUIfield(val,start) {
/*	el = element (x,y,w,h,str,val)
	var = pixels +/- variance
	start = where in mu_ui to start
	returns array of windows control names of next elements in line
*/
	global mu_UI
	qx := []
	el := MorUIfind(val,start)
	var := 3
	
	loop, % mu_UI.MaxIndex()
	{
		if (A_index<start) {
			continue
		}
		i := mu_UI[A_Index]
		if !(i.str~="i)EDIT|Button|COMBOBOX") {
			continue
		}
		if (i.x < el.x+el.w) {
			continue
		}
		if ((i.y>el.y-var) and (i.y<el.y+var)) {
			q .= substr("000" i.x,-3) "- " A_index "`n"
		}
	}
	sort, q
	loop, parse, q, `n, `r
	{
		res := strx(A_LoopField,"- ",1,2,"`n",1)
		qx.push(mu_UI[res].str)
	}
	return qx
}

MorUIfill(start,win) {
/*	fields = array of labels:values to fill
	start = starting line
	win = winID to use
*/
	global ptDem, user
	
	fields := {"ID":ptDem["mrn"]
			,"Last Name":ptDem["nameL"],"First":ptDem["nameF"]
			,"Gender":ptDem["Sex"],"DOB":ptDem["DOB"]
			,"Referring Physician":ptDem["Provider"],"Hookup Tech":user
			,"Indications":RegExReplace(ptDem["Indication"],"\|",";")}
	
	WinActivate, ahk_id %win%
	for key,val in fields
	{
		el := MorUIfield(key,start)
		if (key="DOB") {
			dobEdit := []
			dobCombo := []
			dt := parseDate(val)
			loop, % el.MaxIndex() 
			{
				x := el[A_index]
				if instr(x,"edit") {
					dobEdit.push(x)
				}
				if instr(x,"combobox") {
					dobCombo.push(x)
				}
			}
			uiFieldFill(dobEdit[1],dt.DD,win)
			uiFieldFill(dobCombo[2],dt.MMM,win)
			uiFieldFill(dobEdit[2],dt.YYYY,win)
			continue
		}
		uiFieldFill(el[1],val,win)
	}
	return
}

UiFieldFill(fld,val,win) {
	cb := []
	ControlSetText, % fld, % val, ahk_id %win%
	if instr(fld,"COMBOBOX") {
		ControlGet, cbox, List,, % fld, ahk_id %win%
		loop, parse, cbox, `n, `r
		{
			cb[A_index] := A_LoopField
		}
		Control, Choose, % ObjHasValue(cb,val), % fld, ahk_id %win%
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
	fldval := Object()
	labels := Object()
	blk := Object()
	blk2 := Object()
	ptDem := Object()
	pt := Object()
	chk := Object()
	matchProv := Object()
	fileOut := fileOut1 := fileOut2 := ""
	summBl := summ := ""
	fullDisc := ""
	monType := ""
	
	if (instr(newtxt,"zio xt")) {															; Processing loop based on identifying string in newtxt
		gosub Zio
	} else if (instr(newtxt,"Preventice") && instr(newtxt,"HScribe")) 	{					; New Preventice Holter 2017
		gosub Holter_Pr2
	} else if (instr(newtxt,"Preventice") && instr(newtxt,"End of Service Report")) {		; Body Guardian Heart CEM
		gosub Event_BGH
	;~ } else if ((newtxt~="i)Philips|Lifewatch") && instr(newtxt,"Holter")) {				; Obsolete LW Holters
		;~ gosub Holter_LW
	;~ } else if ((newtxt~="i)Philips|Lifewatch") && InStr(newtxt,"Transmission")) {		; Lifewatch event
		;~ gosub Event_LW
	;~ } else if (instr(newtxt,"Preventice") && instr(newtxt,"H3Plus")) {					; Original Preventice Holter
		;~ gosub Holter_Pr
	} else {
		eventlog(fileNam " bad file.")
		MsgBox No match!
		return
	}
	if (fetchQuit=true) {																	; exited demographics fetchGUI
		return																				; so skip processing this file
	}
	gosub epRead																			; find out which EP is reading today
	
	gosub outputfiles																		; generate and save output CSV, rename and move PDFs
return
}

outputfiles:
{
	/*	Output the results and move files around
	*/
	fileOut1 .= (substr(fileOut1,0,1)="`n") ?: "`n"											; make sure that there is only one `n 
	fileOut2 .= (substr(fileOut2,0,1)="`n") ?: "`n"											; on the header and data lines
	fileout := fileOut1 . fileout2															; concatenate the header and data lines
	tmpDate := parseDate(fldval["Test_Date"])												; get the study date
	filenameOut := fldval["MRN"] " " fldval["Name_L"] " " tmpDate.MM "-" tmpDate.DD "-" tmpDate.YYYY
	tmpFlag := tmpDate.YYYY . tmpDate.MM . tmpDate.DD . "020000"
	
	FileDelete, .\tempfiles\%fileNameOut%.csv												; clear any previous CSV
	FileAppend, %fileOut%, .\tempfiles\%fileNameOut%.csv									; create a new CSV in tempfiles
	
	impSub := (monType~="BGH") ? "Event\" : "Holter\"										; Import subfolder Event or Holter
	FileCopy, .\tempfiles\%fileNameOut%.csv, %importFld%%impSub%*.*, 1						; copy CSV from tempfiles to importFld\impSub
	
	if (FileExist(fileIn "sh.pdf")) {														; filename for OnbaseDir
		fileHIM := fileIn "sh.pdf"															; prefer shortened if it exists
	} else {
		fileHIM := fileIn
	}
	FileCopy, % fileHIM, % OnbaseDir1 filenameOut ".pdf", 1									; Copy to OnbaseDir
	FileCopy, % fileHIM, % OnbaseDir2 filenameOut ".pdf", 1									; Copy to HCClinic folder *** DO WE NEED THIS? ***
	
	FileCopy, %fileIn%, %holterDir%Archive\%filenameOut%.pdf, 1								; Copy the original PDF to holterDir Archive
	FileCopy, %fileIn%sh.pdf, %holterDir%%filenameOut%-short.pdf, 1							; Copy the shortened PDF, if it exists
	FileDelete, %fileIn%																	; Need to use Copy+Delete because if file opened
	FileDelete, %fileIn%sh.pdf																;	was never completing filemove
	FileSetTime, tmpFlag, %holterDir%Archive\%filenameOut%.pdf, C							; set the time of PDF in holterDir to 020000 (processed)
	FileSetTime, tmpFlag, %holterDir%%filenameOut%-short.pdf, C
	eventlog("Move files '" fileIn "' -> '" filenameOut)
	
	fileWQ := ma_date "," user "," 															; date processed and MA user
			. """" chk.Prov """" ","														; extracted provider
			. """" fldval["Name_L"] ", " fldval["Name_F"] """" ","							; CIS name
			. """" fldval["MRN"] """" ","													; CIS MRN
			. """" fldval["dem-Test_date"] """" ","											; extracted Test date (or CIS encounter date if none)
			. """" fldval["dem-Test_end"] """" ","											; extracted Test end
			. """" fldval["dem-Site"] """" ","												; CIS location
			. """" fldval["dem-Indication"] """" ","										; Indication
			. """" monType """" ; ","														; Monitor type
			. "`n"
	FileAppend, %fileWQ%, .\logs\fileWQ.csv													; Add to logs\fileWQ list
	
	moveWQ(fldval["wqid"])																	; Move enroll[@id] from Pending to Done list
	
Return
}

moveWQ(id) {
	global wq, fldval
	
	filecheck()
	FileOpen(".lock", "W")															; Create lock file.
	
	wqStr := "/root/pending/enroll[@id='" id "']"
	x := wq.selectSingleNode(wqStr)
	date := x.selectSingleNode("date").text
	mrn := x.selectSingleNode("mrn").text
	
	if (mrn) {																			; record exists
		wq.addElement("done",wqStr,{user:A_UserName},A_Now)								; set as done
		x := wq.selectSingleNode("/root/pending/enroll[@id='" id "']")					; reload x node
		clone := x.cloneNode(true)
		wq.selectSingleNode("/root/done").appendChild(clone)							; copy x.clone to DONE
		x.parentNode.removeChild(x)														; remove x
		eventlog("wqid " id " (" mrn " from " date ") moved to DONE list.")
	} else {
		id := A_TickCount
		wq.addElement("enroll","/root/done",{id:id})
		newID := "/root/pending/enroll[@id='" id "']"
		wq.addElement("date",newID,fldval["dem-Test_date"])
		wq.addElement("name",newID,fldval["dem-Name"])
		wq.addElement("mrn",newID,fldval["dem-MRN"])
		wq.addElement("done",newID,{user:A_UserName},A_Now)
		eventlog("No wqid. Saved new DONE record " fldval["dem-MRN"] ".")
	}
	wq.save("worklist.xml")
	
	FileDelete, .lock
	
	return
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
	
	RegExMatch(y.selectSingleNode("//call[@date='" dlDate "']/EP").text, "Oi)(Chun|Salerno|Seslar)", ymatch)
	if !(ymatch := ymatch.value()) {
		ymatch := epMon ? epMon : cmsgbox("Electronic Forecast not complete","Which EP on Monday?","Chun|Salerno|Seslar","Q")
		epMon := ymatch
		eventlog("Reading EP assigned to " epMon ".")
	}
	
	if (RegExMatch(fldval["ordering"], "Oi)(Chun|Salerno|Seslar)", epOrder))  {
		ymatch := epOrder.value()
	}
	
	FormatTime, ma_date, A_Now, MM/dd/yyyy
	fileOut1 .= ",""EP_read"",""EP_date"",""MA"",""MA_date"""
	fileOut2 .= ",""" ymatch """,""" niceDate(dlDate) """,""" user """,""" ma_date """"
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
	if !FileExist(fileIn "sh.pdf") {
		FileCopy, %fileIn%, %fileIn%sh.pdf
	}
	FileGetSize, sizeIn, %fileIn%
	FileGetSize, sizeOut, %fileIn%sh.pdf
	eventlog("IN: " thousandsSep(sizeIn) ", OUT: " thousandsSep(sizeOut))
	progress, off
return	
}

Holter_Pr2:
{
	eventlog("Holter_Pr2")
	monType := "PR"
	fullDisc := "i)60\s+s(ec)?/line"
	
	demog := stregX(newtxt,"Name:",1,0,"Conclusions",1)
	
	gosub checkProcPR2											; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	/* Holter PDF is valid. OK to process.
	 * Pulls text between field[n] and field[n+1], place in labels[n] name, with prefix "dem-" etc.
	 */
	fields[1] := ["Name","\R","Recording Start Date/Time","\R"
		, "ID","Secondary ID","Admission ID","\R"
		, "Date Of Birth","Age","Gender","\R"
		, "Date Processed","(Referring|Ordering) Phys(ician)?","\R"
		, "Technician|Hookup Tech","Recording Duration","\R"
		, "Analyst","Recorder (No|Number)","\R"
		, "Indications","Medications","\R"
		, "Hookup time","Location","Acct Num"]
	labels[1] := ["Name","null","Test_date","null"
		, "null","MRN","null","null"
		, "DOB","VOID_Age","Sex","null"
		, "Scan_date","Ordering","null"
		, "Hookup_tech","VOID_Duration","null"
		, "Scanned_by","Device_SN","null"
		, "Indication","VOID_meds","null"
		, "Hookup_time","Site","Billing"]
	fieldvals(demog,1,"dem")
	
	sumStat := RegExReplace(columns(stregX(newtxt,"\s+Scan Criteria",1,0,"\s+RR Variability\s+\(",0)
		,"\s+Summary Statistics","\s+RR Variability",0
		,"VENTRICULAR ECTOPY","SUPRAVENTRICULAR ECTOPY"),": ",":   ")
	
	rateStat := stregX(sumStat,"ALL BEATS",1,0,"VENTRICULAR ECTOPY",1)
	fields[1] := ["Total QRS", "Recording Duration", "Analyzed Duration"
		, "Minimum HR","Maximum HR","Average HR"
		, "Longest Tachycardia","Fastest Tachycardia","Longest Bradycardia","Slowest Bradycardia"
		, "Longest RR", "Pauses .* ms"]
	labels[1] := ["Total_beats", "dem:Recording_time", "dem:Analysis_time"
		, "Min","Max","Avg"
		, "Longest_tachy","Fastest","Longest_brady","Slowest"
		, "sve:LongRR","sve:Pauses"]
	scanParams(rateStat,1,"hrd",1)
	fldVal["dem-Test_end"] := RegExReplace(fldVal["dem-Recording_time"],"(\d{1,2}) hr (\d{1,2}) min","$1:$2")	; Places value for fileWQ, without affecting fileOut
	
	rateStat := stregX(sumStat,"VENTRICULAR ECTOPY",1,0,"PACED|SUPRAVENTRICULAR ECTOPY",1)
	fields[2] := ["Ventricular Beats","Singlets","Couplets","Runs","Fastest Run","Longest Run","R on T Beats"]
	labels[2] := ["Total","SingleVE","Couplets","Runs","Fastest","Longest","R on T"]
	scanParams(rateStat,2,"ve",1)
	
	rateStat := stregX(sumStat "<<<","SUPRAVENTRICULAR ECTOPY",1,0,"<<<|OTHER RHYTHM EPISODES",1)
	fields[3] := ["Supraventricular Beats","Singlets","Pairs","Runs","Fastest Run","Longest Run"]
	labels[3] := ["Total","Single","Pairs","Runs","Fastest","Longest"]
	scanParams(rateStat,3,"sve",1)
	
	LWify()
	tmpstr := stregx(newtxt,"Conclusions",1,1,"Reviewing Physician",1)
	StringReplace, tmpstr, tmpstr, `r, `n, ALL
	fileout1 .= """INTERP"""
	fileout2 .= """" trim(cleanspace(tmpstr)," `n") """"
	fileOut1 .= ",""Mon_type"""
	fileOut2 .= ",""Holter"""
	
	ShortenPDF(fullDisc)

return
}

CheckProcPr2:
{
	eventlog("CheckProcPr")
	chk.Name := strVal(demog,"Name","\R")												; Name
		chk.Name := RegExReplace(chk.Name,"i),?( JR| III| IV)$")							; Filter out 
		chk.Last := trim(strX(chk.Name,"",1,1,",",1,1)," `r`n")								; NameL				must be [A-Z]
		chk.First := trim(strX(chk.Name,",",1,1,"",0)," `r`n")								; NameF				must be [A-Z]
	chk.MRN := strVal(demog,"Secondary ID","Admission ID")								; MRN
	chk.DOB := strVal(demog,"Date of Birth","Age")										; DOB
	chk.Sex := strVal(demog,"Gender","\R")												; Sex
	chk.Prov := cleanspace(strVal(demog,"(Ordering|Referring) Phys(ician)?","\R"))		; Ordering MD
	chk.Ind := strVal(demog,"Indications","Medications")								; Indication
	chk.Med := strVal(demog,"Medications","\R")											; Meds (contains upload code)
	chk.Ser := strVal(demog,"Recorder N(o|umber)","\R")									; Ser Num
	chk.Date := strVal(demog,"Recording Start Date/Time","\R")							; Study date
	
	chkDT := parseDate(chk.Date)
	chkFilename := chk.MRN " * " chkDT.MM "-" chkDT.DD "-" chkDT.YYYY
	if FileExist(holterDir . "Archive\" . chkFilename . ".pdf") {
		FileDelete, %fileIn%
		eventlog(chk.MRN " PDF archive exists, deleting '" fileIn "'")
		fetchQuit := true
		return
	}
	tmpWQ := findWQid(chkDT.YYYY chkDT.MM chkDT.DD,chk.MRN,chk.Name)
	fldval["wqid"] := tmpWQ.id
	pt := readwq(tmpWQ.id)
	if (tmpWQ.node = "done") {
		MsgBox File has been scanned already.
		eventlog(fileIn " already scanned.")
		fetchQuit := true
		return
	}
	if (fileinsize < 3000000) {															; Shortened files are usually < 1-2 Meg
		eventlog("Filesize predicts non-full disclosure PDF.")							; Full disclosure are usually ~ 9-19 Meg
		MsgBox, 4112, Filesize error!, This file does not appear to be a full-disclosure PDF. Please download the proper file and try again.
		fetchQuit := true
		return
	}
	Run , pdftotext.exe "%fileIn%" tempfull.txt,,min,wincons							; convert PDF all pages to txt file
	eventlog("Extracting full text.")
	
	if (pt.acct) {																		; <acct> exists, has been registered or uploaded through TRRIQ
		ptDem["mrn"] := pt.mrn															; fill ptDem[] with values
		ptDem["loc"] := pt.site
		ptDem["EncDate"] := pt.date
		ptDem["Account Number"] := RegExMatch(pt.acct,"([[:alpha:]]+)(\d{8,})",z) ? z2 : pt.acct
		ptDem["nameL"] := strX(pt.name,"",0,1,",",1,1)
		ptDem["nameF"] := strX(pt.name,",",1,1,"",0)
		ptDem["Sex"] := pt.sex
		ptDem["dob"] := pt.dob
		ptDem["Provider"] := pt.prov
		ptDem["Indication"] := pt.ind
		ptDem["loc"] := z1
		eventlog("Pulled valid data for " pt.name " " pt.mrn " " pt.date)
		MsgBox, 4160, Found valid registration, % "" 
		  . pt.Name "`n" 
		  . "MRN " pt.MRN "`n" 
		  . "Acct " pt.Acct "`n" 
		  . "Ordering: " pt.Prov "`n" 
		  . "Study date: " pt.Date "`n`n" 
	} else {																			; no prior TRRIQ data
		eventlog("PDF demog: " chk.MRN " - " chk.Last ", " chk.First)
		Clipboard := chk.Last ", " chk.First											; fill clipboard with name, so can just paste into CIS search bar
		MsgBox, 4096,, % "Extracted data for:`n"
			. "   " chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
			
		ptDem["nameL"] := chk.Last														; Placeholder values for fetchGUI from PDF
		ptDem["nameF"] := chk.First
		ptDem["mrn"] := chk.MRN
		ptDem["DOB"] := chk.DOB
		ptDem["Sex"] := chk.Sex
		ptDem["Loc"] := chk.Loc
		ptDem["Account number"] := chk.Acct												; If want to force click, don't include Acct Num
		ptDem["Provider"] := filterProv(chk.Prov).name
		ptDem["EncDate"] := chk.Date
		ptDem["Indication"] := chk.Ind
		
		fetchQuit:=false
		gosub fetchGUI
		gosub fetchDem
		
		tmp:=fuzzysearch(format("{:U}",chk.Last ", " chk.First), format("{:U}",ptDem["nameL"] ", " ptDem["nameF"]))
		if (tmp > 0.15) {
			eventlog("Name error. "
				. "Parsed """ chk.mrn """, """ chk.Last ", " chk.First """ "
				. "Grabbed """ ptDem["mrn"] """, """ ptDem["nameL"] ", " ptDem["nameF"] """.")
				
			if (chk.MRN=ptDem["mrn"]) {													; correct MRN but bad name match
				MsgBox, 262193, % "Name error (" round((1-tmp)*100,2) "%)"
					, % "Name does not match!`n`n"
					.	"	Parsed:	" chk.Last ", " chk.First "`n"
					.	"	Grabbed:	" ptDem["nameL"] ", " ptDem["nameF"] "`n`n"
					.	"OK = use " ptDem["nameL"] ", " ptDem["nameF"] "`n`n"			; "OK" will accept this fetchDem data
					.	"Cancel = skip this file"
				IfMsgBox, Cancel
				{
					eventlog("Cancel this PDF.")
					fetchQuit:=true														; cancel out of processing file
					return
				}
			} else {																	; just plain doesn't match
				MsgBox, 262160, % "Name error (" round((1-tmp)*100,2) "%)"
					, % "Name does not match!`n`n"
					.	"	Parsed:	" chk.Last ", " chk.First "`n"
					.	"	Grabbed:	" ptDem["nameL"] ", " ptDem["nameF"] "`n`n"
					.	"Skipping this file."
					
				eventlog("Demographics mismatch.")
				fetchQuit:=true
				return
			}
		}
	}
	/*	When fetchDem successfully completes,
	 *	replace the fields in demog with newly acquired values
	 */
	chk.Name := ptDem["nameL"] ", " ptDem["nameF"] 
	fldval["name_L"] := ptDem["nameL"]
	fldval["name_F"] := ptDem["nameF"]
	demog := RegExReplace(demog,"i`a)Name: (.*)\R","Name:   " chk.Name "   `n")
	demog := RegExReplace(demog,"i)Secondary ID: (.*) Admission ID:","Secondary ID:   " ptDem["mrn"] "                   Admission ID:")
	demog := RegExReplace(demog,"i)Date Of Birth: (.*) Age:", "Date Of Birth:   " ptDem["DOB"] "  Age:")
	demog := RegExReplace(demog,"i`a)(Ordering|Referring) Phys(ician)?:? (.*)\R", "Referring Physician:   " ptDem["Provider"] "`n")
	demog := RegExReplace(demog,"i`a)Indications: (.*) Medications:", "Indications:   " ptDem["Indication"] "   Medications:")	
	demog := RegExReplace(demog,"i`a)Recording Start Date/Time: (.*)\R", "Recording Start Date/Time:   " chk.Date "`n")
	demog := RegExReplace(demog,"i`a)Analyst: (.*) Recorder N(o|umber)","Analyst:   $1   Recorder No")
	demog := RegExReplace(demog,"i`a)Technician: (.*) Recording Duration","Hookup Tech:   $1   Recording Duration")
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
	
	RunWait, pdftotext.exe -table -fixed 3 "%fileIn%" temp.txt, , hide				; reconvert entire Zio PDF 
	newTxt:=""																		; clear the full txt variable
	FileRead, maintxt, temp.txt														; load into maintxt
	StringReplace, newtxt, maintxt, `r`n`r`n, `r`n, All
	StringReplace, newtxt, newtxt, % chr(12), >>>page>>>`r`n, All
	FileDelete tempfile.txt															; remove any leftover tempfile
	FileAppend %newtxt%, tempfile.txt												; create new tempfile with newtxt result
	FileMove tempfile.txt, .\tempfiles\%fileNam%.txt, 1								; overwrite copy in tempfiles
	eventlog("Zio PDF rescanned -> " fileNam ".txt")
	
	zcol := columns(newtxt,"","SIGNATURE",0,"Enrollment Period") ">>>end"
	demo1 := onecol(cleanblank(stregX(zcol,"\s+Date of Birth",1,0,"Prescribing Clinician",1)))
	demo2 := onecol(cleanblank(stregX(zcol,"\s+Prescribing Clinician",1,0,"\s+(Supraventricular Tachycardia \(|Ventricular tachycardia \(|AV Block \(|Pauses \(|Atrial Fibrillation)",1)))
	demog := RegExReplace(demo1 "`n" demo2,">>>end") ">>>end"
	
	gosub checkProcZio											; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	znam := strVal(demog,"Name","Date of Birth")
	formatField("dem","Name_L",strX(znam, "", 1,0, ",", 1,1))
	formatField("dem","Name_F",strX(znam, ", ", 1,2, "", 0))
	
	fields[1] := ["Date of Birth","Patient ID","Gender","Primary Indication","Prescribing Clinician","Managing Location",">>>end"]
	labels[1] := ["DOB","MRN","Sex","Indication","Ordering","Site","end"]
	fieldvals(demog,1,"dem")
	
	formatField("dem","Test_date",fldval["Test_date"])
	formatField("dem","Billing",fldval["Acct"])
	
	znums := columns(zcol ">>>end","Enrollment Period",">>>end",1)
	
	formatField("dem","Recording_time",chk.enroll)
	formatField("dem","Analysis_time",chk.Analysis)
	
	zrate := columns(znums,"Heart Rate","Patient Events",1)
	fields[3] := ["Max","Min","Avg","\R"]
	labels[3] := ["Max","Min","Avg","null"]
	fieldvals(zrate,3,"hrd")
	
	zevent := columns(znums,"Patient Events","Ectopics",1) ">>>end"
	zev_T := columns(zevent,"Triggered","(Diary|>>>end)",1,"Findings")
	fields[4] := ["Events","Findings within(.*)Triggers"]
	labels[4] := ["Triggers","Trigger_Findings"]
	fieldvals(zev_T,4,"event")
	
	zev_D := columns(zevent,"Diary","(Triggered|>>>end)",1,"Findings")
	fields[5] := ["Entries","Findings within(.*)Entries"]
	labels[5] := ["Diary","Diary_Findings"]
	fieldvals(zev_D,5,"event")

	zectopics := columns(znums ">>>end","Ectopics",">>>end",0) ">>>end"
	zsve := columns(zectopics,"Supraventricular Ectopy \(SVE/PACs\)","Ventricular Ectopy \(VE/PVCs\)",1)
	fields[7] := ["Isolated","Couplet","Triplet","\R"]
	labels[7] := ["Single","Pairs","Triplets","null"]
	fieldvals(zsve,7,"sve")
	zsve_tot := (fldval["sve-Single"] ? fldval["sve-Single"] : 0) 
				+ 2*(fldval["sve-Pairs"] ? fldval["sve-Pairs"] : 0) 
				+ 3*(fldval["sve-Triplets"] ? fldval["sve-Triplets"] : 0)
	formatField("sve","Total",zsve_tot)

	if (zsve := stregX(newtxt,"Episode Heart Rates.* SVT",1,0,">>>page>>>",1)) {
		zsve := stregX(zsve,"^(.*)SVT with",1,0,"^(.*)Patient:",1) ">>>>>"
		zsve_fastest := stregX(zsve,"^(.*)Fastest Heart Rate",1,0,"^(.*)Fastest Avg",1) ">>>>>"
		zsve_tmp := columns(zsve_fastest,"",">>>>>",,"Average") ">>>>>"
		zsve_fastest := columns(zsve_tmp,"","Average",,"# Beats","Duration")
						. columns(zsve_tmp,"Average",">>>>>",,"Range","Pt Triggered")
		fields[7] := ["Fastest Heart Rate","# Beats","Duration","Average","Range","Pt Triggered"]
		labels[7] := ["Fastest_time","Beats","null","null","Fastest","null"]
		fieldvals(zsve_fastest,7,"sve")
		zsve_longest := stregX(zsve,"^(.*)Longest SVT",1,0,">>>>>",0)
		zsve_tmp := columns(zsve_longest,"",">>>>>",,"Average") ">>>>>"
		zsve_longest := columns(zsve_tmp,"","Average",,"# Beats","Duration")
						. columns(zsve_tmp,"Average",">>>>>",,"Range","Pt Triggered")
		fields[7] := ["Longest SVT Episode","# Beats","Duration","Average","Range","Pt Triggered"]
		labels[7] := ["Longest_time","Longest","null","null","null","null"]
		fieldvals(zsve_longest,7,"sve")
	}
	
	zve := columns(zectopics,"Ventricular Ectopy \(VE/PVCs\)",">>>end",1)
	fields[8] := ["Isolated","Couplet","Triplet","Longest Ventricular Bigeminy Episode","Longest Ventricular Trigeminy Episode"]
	labels[8] := ["SinglePVC","Couplets","Triplets","LongestBigem","LongestTrigem"]
	fieldvals(zve,8,"ve")
	zve_tot := (fldval["ve-SinglePVC"] ? fldval["ve-SinglePVC"] : 0) 
				+ 2*(fldval["ve-Couplets"] ? fldval["ve-Couplets"] : 0) 
				+ 3*(fldval["ve-Triplets"] ? fldval["ve-Triplets"] : 0)
	formatField("ve","Total",zve_tot)
	
	if (zve := stregX(newtxt,"Episode Heart Rates.* VT",1,0,">>>page>>>",1)) {
		zve := stregX(zve,"^(.*)VT with",1,0,"^(.*)Patient:",1) ">>>>>"
		zve_fastest := stregX(zve,"^(.*)Fastest Heart Rate",1,0,"^(.*)Fastest Avg",1) ">>>>>"
		zve_tmp := columns(zve_fastest,"",">>>>>",,"Average") ">>>>>"
		zve_fastest := columns(zve_tmp,"","Average",,"# Beats","Duration")
						. columns(zve_tmp,"Average",">>>>>",,"Range","Pt Triggered")
		fields[8] := ["Fastest Heart Rate","# Beats","Duration","Average","Range","Pt Triggered"]
		labels[8] := ["Fastest_time","Beats","null","null","Fastest","null"]
		fieldvals(zve_fastest,8,"ve")
		zve_longest := stregX(zve,"^(.*)Longest VT",1,0,">>>>>",0)
		zve_tmp := columns(zve_longest,"",">>>>>",,"Average") ">>>>>"
		zve_longest := columns(zve_tmp,"","Average",,"# Beats","Duration")
						. columns(zve_tmp,"Average",">>>>>",,"Range","Pt Triggered")
		fields[8] := ["Longest VT Episode","# Beats","Duration","Average","Range","Pt Triggered"]
		labels[8] := ["Longest_time","Longest","null","null","null","null"]
		fieldvals(zve_longest,8,"ve")
	}
	
	LWify()
	zinterp := cleanspace(columns(newtxt,"Preliminary Findings","SIGNATURE",,"Final Interpretation"))
	zinterp := trim(StrX(zinterp,"",1,0,"Final Interpretation",1,20))
	fileout1 .= """INTERP"""
	fileout2 .= """" . zinterp . """"
	fileOut1 .= ",""Mon_type"""
	fileOut2 .= ",""Holter"""
	
	FileCopy, %fileIn%, %fileIn%sh.pdf

return
}

ZioArrField(txt,fld) {
	str := stregX(txt,fld,1,0,"#####",1)
	if instr(str,"Episodes") {
		;~ str := strX(columns(str,fld,"#####",0,"Episodes"),"Episodes",1,0,"",0)
		str := columns(str,fld,"#####",0,"Episodes")
		str := RegExReplace(str,"i)None found")
	}
	Loop, parse, str, `n,`r
	{
		i:=A_LoopField
		if (i~=fld) {															; skip header line
			continue
		}
		if !(trim(i)) {                                    						; skip entirely blank lines 
			continue
		}
		newStr .= i "`n"   							                           ; only add lines with text in it 
	} 
	if (newStr="") {
		newStr := "None found`n"
	}
	return trim(cleanspace(newStr))
}

CheckProcZio:
{
	chk.Name := trim(cleanSpace(stregX(zcol,"Report for",1,1,"Date of Birth",1)))
		chk.Name := RegExReplace(chk.Name,"i),?( JR| III| IV)$")									; Filter out 
		chk.Last := trim(strX(chk.Name, "", 1,1, ",", 1,1))
		chk.First := trim(strX(chk.Name, ", ", 1,2, "", 0))
	chk.DOB := RegExReplace(strVal(demog,"Date of Birth","Patient ID"),"\s+\(.*(yrs|mos)\)")		; DOB
	chk.MRN := strVal(demog,"Patient ID","Gender")											; MRN
	chk.Sex := strVal(demog,"Gender","Primary Indication")											; Sex
	chk.Ind := RegExReplace(strVal(demog,"Primary Indication","Prescribing Clinician"),"\(R00.0\)\s+")				; Indication
	chk.Prov:= strVal(demog,"Prescribing Clinician","(Referring Clinician|Managing Location)")												; Ordering MD
	chk.Loc := strVal(demog,"Managing Location",">>>end")											; MRN
	
	demog := "Name   " chk.Last ", " chk.First "`n" demog
	
	tmp := oneCol(stregX(zcol,"Enrollment Period",1,0,"Heart\s+Rate",1))
		chk.enroll := strVal(tmp,"Enrollment Period","Analysis Time")
		chk.DateStart := strVal(chk.enroll,"hours?",",")
		chk.DateEnd := strVal(chk.enroll,"to\s",",")
		chk.Date := chk.DateStart
		chk.Analysis := strVal(tmp,"Analysis Time","\(after")
		chk.enroll := stregX(chk.enroll,"",1,0,"\s{3}",1)
	
	zcol := stregx(zcol,"\s+(Supra)?ventricular tachycardia \(",1,0,">>>end",1)
	
	/*	
	 *	Return from CheckProc for testing
	 */
		;~ Return
	
	Clipboard := chk.Last ", " chk.First												; fill clipboard with name, so can just paste into CIS search bar
	;~ if (!(chk.Last~="[a-z]+")															; Check field values to see if proper demographics
		;~ && !(chk.First~="[a-z]+") 														; meaning names in ALL CAPS
		;~ && (chk.Acct~="\d{8}"))															; and EncNum present
	;~ {
		;~ MsgBox, 4132, Valid PDF, % ""
			;~ . chk.Last ", " chk.First "`n"
			;~ . "MRN " chk.MRN "`n"
			;~ . "Acct " chk.Acct "`n"
			;~ . "Ordering: " chk.Prov "`n"
			;~ . "Study date: " chk.DateStart "`n`n"
			;~ . "Is all the information correct?`n"
			;~ . "If NO, reacquire demographics."
		;~ IfMsgBox, Yes																; All tests valid
		;~ {
			;~ return																	; Select YES, return to processing Holter
		;~ } 
		;~ else 																		; Select NO, reacquire demographics
		;~ {
			;~ MsgBox, 4096, Adjust demographics, % chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			;~ . "Paste clipboard into CIS search to select patient and encounter"
		;~ }
	;~ }
	;~ else 																			; Not valid PDF, get demographics post hoc
	{
		eventlog("PDF demog: " chk.MRN " - " chk.Last ", " chk.First)
		MsgBox, 4096,, % "Extracted data for:`n   " chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
	}
	; Either invalid PDF or want to correct values
	ptDem["nameL"] := chk.Last															; Placeholder values for fetchGUI from PDF
	ptDem["nameF"] := chk.First
	ptDem["MRN"] := chk.MRN
	ptDem["DOB"] := chk.DOB
	ptDem["Sex"] := chk.Sex
	ptDem["Loc"] := chk.Loc
	ptDem["Account number"] := chk.Acct													; If want to force click, don't include Acct Num
	ptDem["Provider"] := filterProv(chk.Prov).name
	ptDem["EncDate"] := chk.DateStart
	ptDem["Indication"] := chk.Ind
	
	fetchQuit:=false
	gosub fetchGUI
	gosub fetchDem
	
	if (tmp:=fuzzysearch(chk.Last " " chk.First, ptDem["nameL"] " " ptDem["nameF"]) > 0.15) {
		MsgBox, 262160, % "Name error (" round((1-tmp)*100,2) "%)"
			, % "Name does not match!`n`n"
			.	"	Parsed:	" chk.Last ", " chk.First "`n"
			.	"	Grabbed:	" ptDem["nameL"] ", " ptDem["nameF"] "`n`n"
			.	"Skipping this file."
			
		eventlog("Name error. Parsed """ chk.Last ", " chk.First """ Grabbed """ ptDem["nameL"] ", " ptDem["nameF"] """.")
		fetchQuit:=true
		return
	}
	/*	When fetchDem successfully completes,
	 *	replace the fields in demog with newly acquired values
	 */
	fldval["Test_date"] := chk.DateStart
	fldval["Test_end"] := chk.DateEnd
	fldval["name_L"] := ptDem["nameL"]
	fldval["name_F"] := ptDem["nameF"]
	fldval["MRN"] := ptDem["MRN"]
	fldval["Acct"] := ptDem["Account Number"]
	fldval["dem-Billing"] := fldval["Acct"]
	fldval["dem-Test_date"] := chk.DateStart
	fldval["dem-Test_end"] := chk.DateEnd
	
	demog := RegExReplace(demog,"i)Name(.*)Date of Birth","Name   " ptDem["nameL"] ", " ptDem["nameF"] "`nDate of Birth",,1)
	demog := RegExReplace(demog,"i)Date of Birth(.*)Patient ID","Date of Birth   " ptDem["DOB"] "`nPatient ID",,1)
	demog := RegExReplace(demog,"i)Patient ID(.*)Gender","Patient ID   " ptDem["MRN"] "`nGender",,1)
	demog := RegExReplace(demog,"i)Gender(.*)Primary Indication","Gender   " ptDem["Sex"] "`nPrimary Indication",,1)
	demog := RegExReplace(demog,"i)Primary Indication(.*)Prescribing Clinician","Primary Indication   " ptDem["Indication"] "`nPrescribing Clinician",,1)
	demog := RegExReplace(demog,"i)Prescribing Clinician(.*)(Referring Clinician|Managing Location)","Prescribing Clinician   " ptDem["Provider"] "`nManaging Location",,1)
	demog := RegExReplace(demog,"i)Managing Location(.*)>>>end","Managing Location   " ptDem["loc"] "`n>>>end",,1)
	
	return
}

Event_BGH:
{
	eventlog("Event_BGH")
	monType := "BGH"
	
	name := "Patient Name:   " trim(columns(newtxt,"Patient:","Enrollment Info",1,"")," `n")
	demog := columns(newtxt,"","(Summarized Findings|Event Summary)",,"Enrollment Info")
	enroll := RegExReplace(strX(demog,"Enrollment Info",1,0,"",0),": ",":   ")
	diag := "Diagnosis:   " trim(stRegX(demog,"`a)Diagnosis \(.*\R",1,1,"(Preventice|Enrollment Info)",1)," `n")
	demog := columns(demog,"\s+Patient ID","Diagnosis \(",,"Monitor   ") "#####"
	mon := stregX(demog,"Monitor\s{3}",1,0,"#####",1)
	demog := columns(demog,"\s+Patient ID","Monitor   ",,"Gender","Date of Birth","Phone")		; columns get stuck in permanent loop
	demog := name "`n" demog "`n" mon "`n" diag "`n"
	
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
	
	fields[2] := ["Date Recorded","Date Ended","\R"]
	labels[2] := ["Test_date","Test_end","VOID"]
	fieldvals(enroll,2,"dem")
	fieldColAdd("dem","Billing",ptDem["Account Number"])
	
	fields[3] := ["Critical","Total","Serious","(Manual|Pt Trigger)","Stable","Auto Trigger"]
	labels[3] := ["Critical","Total","Serious","Manual","Stable","Auto"]
	fieldvals(enroll,3,"counts")
	
	fileOut1 .= ",""Mon_type"""
	fileOut2 .= ",""Event"""
	
	FileCopy, %fileIn%, %fileIn%sh.pdf
	
Return
}

CheckProcBGH:
{
	chk.Name := strVal(demog,"Patient Name","Patient ID")										; Name
		chk.Name := RegExReplace(chk.Name,"i),?( JR| III| IV)$")									; Filter out JR
		chk.First := trim(strX(chk.Name,"",1,1," ",1,1)," `r`n")									; NameL				must be [A-Z]
		chk.Last := trim(strX(chk.Name," ",0,1,"",0)," `r`n")										; NameF				must be [A-Z]
	chk.MRN := strVal(demog,"Patient ID","Physician")											; MRN
	chk.Prov := strVal(demog,"Physician","Gender")												; Ordering MD
	chk.Sex := strVal(demog,"Gender","Date of Birth")											; Sex
	chk.DOB := strVal(demog,"Date of Birth","Practice")											; DOB
	chk.Ind := strVal(demog,"Diagnosis","\R")													; Indication
	chk.Date := strVal(enroll,"Period \(.*\)","Event Counts")									; Study date
		chk.DateEnd := trim(strX(chk.Date," - ",0,3,"",0)," `r`n")
		chk.DateStart := trim(strX(chk.Date,"",1,1," ",1,1)," `r`n")
	chkDT := parseDate(chk.DateStart)
	fldval["wqid"] := findWQid(chkDT.YYYY chkDT.MM chkDT.DD,chk.MRN,chk.Name).id

	Clipboard := chk.Last ", " chk.First												; fill clipboard with name, so can just paste into CIS search bar
	;~ if (!(chk.Last~="[a-z]+")															; Check field values to see if proper demographics
		;~ && !(chk.First~="[a-z]+") 														; meaning names in ALL CAPS
		;~ && (chk.Acct~="\d{8}"))															; and EncNum present
	;~ {
		;~ MsgBox, 4132, Valid PDF, % ""
			;~ . chk.Last ", " chk.First "`n"
			;~ . "MRN " chk.MRN "`n"
			;~ . "Acct " chk.Acct "`n"
			;~ . "Ordering: " chk.Prov "`n"
			;~ . "Study date: " chk.DateStart "`n`n"
			;~ . "Is all the information correct?`n"
			;~ . "If NO, reacquire demographics."
		;~ IfMsgBox, Yes																; All tests valid
		;~ {
			;~ return																	; Select YES, return to processing Holter
		;~ } 
		;~ else 																		; Select NO, reacquire demographics
		;~ {
			;~ MsgBox, 4096, Adjust demographics, % chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
			;~ . "Paste clipboard into CIS search to select patient and encounter"
		;~ }
	;~ }
	;~ else 																			; Not valid PDF, get demographics post hoc
	{
		eventlog("PDF demog: " chk.MRN " - " chk.Last ", " chk.First)
		MsgBox, 4096,, % "Extracted data for:`n   " chk.Last ", " chk.First "`n   " chk.MRN "`n   " chk.Loc "`n   " chk.Acct "`n`n"
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
	ptDem["Provider"] := filterProv(chk.Prov).name
	ptDem["EncDate"] := chk.DateStart
	ptDem["EndDate"] := chk.DateEnd
	ptDem["Indication"] := chk.Ind
	
	fetchQuit:=false
	gosub fetchGUI
	gosub fetchDem
	
	tmp:=fuzzysearch(format("{:U}",chk.Last ", " chk.First), format("{:U}",ptDem["nameL"] ", " ptDem["nameF"]))
	if (tmp > 0.15) {
		eventlog("Name error. Parsed """ chk.mrn """, """ chk.Last ", " chk.First """ Grabbed """ ptDem["mrn"] """, """ ptDem["nameL"] ", " ptDem["nameF"] """.")
		if (chk.MRN=ptDem["mrn"]) {
			MsgBox, 262193, % "Name error (" round((1-tmp)*100,2) "%)"
				, % "Name does not match!`n`n"
				.	"	Parsed:	" chk.Last ", " chk.First "`n"
				.	"	Grabbed:	" ptDem["nameL"] ", " ptDem["nameF"] "`n`n"
				.	"OK = use " ptDem["nameL"] ", " ptDem["nameF"] "`n`n"
				.	"Cancel = skip this file"
			IfMsgBox, Cancel
			{
				fetchQuit:=true
				return
			}
		} else {
			MsgBox, 262160, % "Name error (" round((1-tmp)*100,2) "%)"
				, % "Name does not match!`n`n"
				.	"	Parsed:	" chk.Last ", " chk.First "`n"
				.	"	Grabbed:	" ptDem["nameL"] ", " ptDem["nameF"] "`n`n"
				.	"Skipping this file."
				
			fetchQuit:=true
			return
		}
	}
	
	/*	When fetchDem successfully completes,
	 *	replace the fields in demog with newly acquired values
	 */
	fldval["dem-Site"] := ptDem["Loc"]
	fldval["dem-Billing"] := ptDem["Account Number"]
	chk.Name := ptDem["nameF"] " " ptDem["nameL"] 
		fldval["name_L"] := ptDem["nameL"]
		fldval["name_F"] := ptDem["nameF"]
	demog := RegExReplace(demog,"i)Patient Name: (.*?)Patient ID","Patient Name:   " chk.Name "`nPatient ID")
	demog := RegExReplace(demog,"i)Patient ID(.*?)Physician","Patient ID   " ptDem["mrn"] "`nPhysician")
	demog := RegExReplace(demog,"i)Physician(.*?)Gender", "Physician   " ptDem["Provider"] "`nGender")
	demog := RegExReplace(demog,"i)Gender(.*?)Date of Birth", "Gender   " ptDem["Sex"] "`nDate of Birth")
	demog := RegExReplace(demog,"i)Date of Birth(.*?)Practice", "Date of Birth   " ptDem["DOB"] "`nPractice")	
	enroll := RegExReplace(enroll,"i)Period(.*?)\R", "$1`nDate Recorded:   " chk.DateStart "`nDate Ended:   " chk.DateEnd "`n") 
	eventlog("Demog replaced.")
	
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
	blk1	= leading regex string to start block
	blk2	= ending regex string to end block
	incl	= if null, include blk1 string; if !null, remove blk1 string
	col2	= string demarcates start of COLUMN 2
	col3	= string demarcates start of COLUMN 3
	col4	= string demarcates start of COLUMN 4
*/
	blk1 := rxFix(blk1,"O",1)													; Adds "O)" to blk1
	blk2 := rxFix(blk2,"O",1)
	RegExMatch(x,blk1,blo1)														; Creates blo1 object out of blk1 match in x
	RegExMatch(x,blk2,blo2)														; *** DO I EVEN USE BLO1 ANYMORE? ***
	
	txt := stRegX(x,blk1,1,((incl) ? 1 : 0),blk2,1)								; Get string between BLK1 and BLK2, with or without INCL bit
	;~ MsgBox % txt
	col2 := RegExReplace(col2,"\s+","\s+")										; Make col search strings more flexible for whitespace
	col3 := RegExReplace(col3,"\s+","\s+")
	col4 := RegExReplace(col4,"\s+","\s+")
	
	loop, parse, txt, `n,`r														; find position of columns 2, 3, and 4
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
		if !(trim(i)) {                        						           ; discard entirely blank lines 
		  continue
		}
		txt1 .= substr(i,1,pos2-1) . "`n"										; TXT1 is from 1 to pos2
		if (col4) {																; Handle the 4 col condition
			pos4ck := pos4														; check start of col4
			while !(substr(i,pos4ck-1,1)=" ") {									
				pos4ck := pos4ck-1
			}
			txt4 .= substr(i,pos4ck) . "`n"
			txt3 .= substr(i,pos3,pos4ck-pos3) . "`n"
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"
			continue
		} 
		if (col3) {																; Handle the 3 col condition
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"
			txt3 .= substr(i,pos3) . "`n"
			continue
		}
		txt2 .= substr(i,pos2) . "`n"											; Remaining is just pos2 to end
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
	bl	= which FIELD block to use
	bl2	= label prefix
*/
	global fields, labels, fldval
	
	for k, i in fields[bl]																; Step through each val "i" from fields[bl,k]
	{
		pre := bl2
		j := fields[bl][k+1]															; Next field [k+1]
		m := (j) 
			?	strVal(x,i,j,n,n)														; ...is not null ==> returns value between
			:	trim(strX(SubStr(x,n),":",1,1,"",0)," `n")								; ...is null ==> returns from field[k] to end
		lbl := labels[bl][A_index]
		if (lbl~="^\w{3}:") {															; has prefix e.g. "dem:name2"
			pre := substr(lbl,1,3)														; change pre for this loop, e.g. "dem"
			lbl := substr(lbl,5)														; change lbl for this loop, e.g. "name2"
		}
		cleanSpace(m)
		cleanColon(m)
		fldval[pre "-" lbl] := m
		fldval[lbl] := m
		
		formatField(pre,lbl,m)
	}
}

strVal(hay,n1,n2,BO:="",ByRef N:="") {
/*	hay = search haystack
	n1	= needle1 begin string
	n2	= needle2 end string
	N	= return end position
*/
	opt := "Oi)"
	RegExMatch(hay,opt . n1 . ":?(?P<res>.*?)" . n2, str, (BO)?BO:1)
	N := str.pos("res")+str.len("res")
	
	if (str.pos("res")=="") {															; RexExMatch fail on n1 or n2 (i.e. bad field needles)
		eventlog("*** strVal fail: ''" n1 "' ... '" n2 "'")								; Note the bad fields
	}

	return trim(str.value("res")," :`n")
}

scanParams(txt,blk,pre:="par",rx:="") {
	global fields, labels, fldval
	colstr = (?<=(\s{2}))(\>\s*)?[^\s].*?(?=(\s{2}))
	Loop, parse, txt, `n,`r
	{
		i := trim(A_LoopField) "  "
		set := trim(strX(i,"",1,0,"  ",1,2)," :")								; Get leftmost column to first "  "
		val := objHasValue(fields[blk],set,rx)
		if !(val) {
			continue
		}
		lbl := labels[blk,val]
		if (lbl~="^\w{3}:") {											; has prefix e.g. "dem:"
			pre0 := substr(lbl,1,3)
			lbl := substr(lbl,5)
		} else {
			pre0 := pre
			lbl := lbl
		}
		
		RegExMatch(i															; Add "  " to end of scan string
				,"O)" colstr													; Search "  text  " as each column 
				,col1)															; return result in var "col1"
		RegExMatch(i
				,"O)" colstr
				,col2
				,col1.pos()+1)
		
		res := col1.value()
		if (col2.value()~="^(\>\s*)(?=[^\s])") {
			res := RegExReplace(col2.value(),"^(\>\s*)(?=[^\s])") " (changed from " col1.value() ")"
		}
		if (col2.value()~="(Monitor.*|\d{2}J.*)") {
			res .= ", Rx " cleanSpace(col2.value())
		}
			
		;~ MsgBox % pre "-" labels[blk,val] ": " res
		fldfill(pre0 "-" lbl, res)
		
		formatfield(pre0,lbl,res)
	}
	return
}

fldfill(var,val) {
/*	Nondestructively fill fields
	If val is empty, return
	Otherwise populate with new value
*/
	global fldval
	
	if (val=="") {																; val is null
		return																	; do nothing
	}
	
	fldval[var] := trim(val," `t`r`n")											; set var as val
	
return
}

rxFix(hay,req,spc:="") {
/*	rxFix
	in	= input string, may or may or not include "Oim)" option modifiers
	req	= required modifiers to output
	spc	= replace spaces
*/
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
	BS := RegExReplace(BS,"\s+","\s+")												; Replace each \s with \s+ so no affected by variable spaces
	ES := RegExReplace(ES,"\s+","\s+")
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
	if (!IsObject(bPat) or !IsObject(ePat)) {
		return error
	}
	
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
	global monType, Docs, ptDem, fldval
	if (txt ~= "\d{1,2} hr \d{1,2} min") {
		StringReplace, txt, txt, %A_Space%hr%A_space% , :
		StringReplace, txt, txt, %A_Space%min , 
	}
	txt:=RegExReplace(txt,"i)BPM|Event(s)?|Beat(s)?|( sec(s)?)")			; 	Remove units from numbers
	txt:=RegExReplace(txt,"(:\d{2}?)(AM|PM)","$1 $2")						;	Fix time strings without space before AM|PM
	txt := trim(txt)
	
	if (lab~="Referring|Ordering") {
		tmpCrd := checkCrd(RegExReplace(txt,"i)^Dr(\.)?\s"))				;	Get Crd, Grp, and Eml via checkCrd() <== shouldn't this already be determined?
		fieldColAdd(pre,lab,tmpCrd.best)
		fieldColAdd(pre,lab "_grp",tmpCrd.group)
		fieldColAdd(pre,lab "_eml",Docs[tmpCrd.Group ".eml",ObjHasValue(Docs[tmpCrd.Group],tmpCrd.best)])
		if (tmpCrd="") {
			eventlog("*** Blank Crd value ***")
		}
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
		if (lab="Name") {																; Break name into Last and First
			fieldColAdd(pre,"Name_L",trim(strX(txt,"",1,0,",",1,1)))
			fieldColAdd(pre,"Name_F",trim(strX(txt,",",1,1,"",0)))
			return
		}
		if (lab="Test_date") {															; Only take Test_date to first " "
			txt := strX(txt,"",1,0," ",1,1)
		}
		if (RegExMatch(txt,"O)^(\d{1,2})\s+hr,\s+(\d{1,2})\s+min",tx)) {				; Convert "x hr, yy mins" to "0x:yy"
			fieldColAdd(pre,lab,zDigit(tx.value(1)) ":" zDigit(tx.value(2)))
			return
		}
		if (lab ~= "(Analysis|Recording)_time") {										; Adjust Analysis_time and Recording_time if misreported as 48:00 Holter
			tmp := strX(txt,"",1,0,":",1,1)
			if (tmp > 36) {																; Greater than "36 hr" recording,
				tmp := zDigit(tmp-24)													; subtract 24 hrs
				txt := RegExReplace(txt,"\d{2}:",tmp ":")
			}
		}
		if (txt ~= "^([0-9.]+( BPM( Avg)?)?).+at.+(\d{1,2}:\d{2}:\d{2}).*(AM|PM)?$") {		;	Split timed results "139 at 8:31:47 AM" into two fields
			tx1 := trim(stregX(txt,"",1,0," at ",1))
			tx2 := trim(stregX(txt "<<<"," at ",1,1,"<<<",1))
			fieldColAdd(pre,lab,tx1)
			fieldColAdd(pre,lab "_time",tx2)
			return
		}
		if (txt = "--- at ---") {
			fieldColAdd(pre,lab,"")
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
			fldval[pre "-" lab] := tx2
			fieldColAdd(pre,lab,tx2)
			return
		}
		if (txt ~= "^[0-9.]+\s+\d{1,2}:\d{2}") {						;	Split timed results "139  8:31AM" into two fields
			tx1 := trim(stregX(txt,"",1,0,"\d{1,2}:\d{2}",1,n))
			tx2 := trim(stregX(txt "<<<","\d{1,2}:\d{2}",1,0,"<<<",1))
			fieldColAdd(pre,lab,tx1)
			fieldColAdd(pre,lab "_time",tx2)
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
	fuzz := 1																			; Initially, fuzz is 100%
	if (x="") {																			; fuzzysearch fails if x = ""
		return 
	}
	for rowidx,row in Docs
	{
		if (substr(rowIdx,-3)=".eml")
			continue
		for colidx,item in row
		{
			if (item="") {																; empty field will break fuzzysearch
				continue
			}
			res := fuzzysearch(x,item)
			if (res<fuzz) {																; less fuzzy, new best match
				fuzz := res
				best:=item
				group:=rowidx
			}
		}
	}
	return {"fuzz":fuzz,"best":best,"group":group}
}

filterProv(x) {
	global sites, sites0
	
	allsites := sites "|" sites0
	RegExMatch(x,"i)-(" allsites ")\s*,",site)
	x := trim(x)																		; trim leading and trailing spaces
	x := RegExReplace(x,"i)^Dr(\.)?(\s)?")												; remove preceding "(Dr. )Veronica..."
	x := RegExReplace(x,"i)^[a-z](\.)?\s")												; remove preceding "(P. )Ruggerie, Dennis"
	x := RegExReplace(x,"i)\s[a-z](\.)?$")												; remove trailing "Ruggerie, Dennis( P.)"
	x := RegExReplace(x,"i)-(" allsites ")\s*,",",")									; remove "SCHMER(-YAKIMA), VERONICA"
	x := RegExReplace(x,"i) (MD|DO)$")													; remove trailing "( MD)"
	x := RegExReplace(x,"i) (MD|DO),",",")												; replace "Ruggerie MD, Dennis" with "Ruggerie, Dennis"
	StringUpper,x,x,T																	; convert "RUGGERIE, DENNIS" to "Ruggerie, Dennis"
	;~ StringUpper,site1,site1,T
	return {name:x, site:site1}
}

httpComm(url:="",verb:="") {
	global servFold
	if (url="") {
		url := "https://depts.washington.edu/pedcards/change/direct.php?" 
				. ((servFold="testlist") ? "test=true&" : "") 
				. "do=" . verb
	}
	whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")							; initialize http request in object whr
		whr.Open("GET"															; set the http verb to GET file "change"
			, url
			, true)
		whr.Send()																; SEND the command to the address
		whr.WaitForResponse()													; and wait for
	return whr.ResponseText														; the http response
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
		if (rx) {
			if RegExMatch(aValue,val) {
				return, key, Errorlevel := 0
			}
		} else {
			if (val = aValue) {
				return, key, ErrorLevel := 0
			}
		}
    return, false, errorlevel := 1
}

strQ(var1,txt) {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
*/
	if (var1="") {
		return error
	}
	return RegExReplace(txt,"###",var1)
}

countlines(hay,n) {
	hay := substr(hay,1,n)
	loop, parse, hay, `n, `r
	{
		max := A_Index
	}
	return max
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

ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-_/]"
	date := []
	time := []
	x := RegExReplace(x,"[,\(\)]")
	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",d) {			; 03-Jan-2015
		date.dd := zdigit(d1)
		date.mmm := d2
		date.mm := zdigit(objhasvalue(mo,d2))
		date.yyyy := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"i)(" moStr "|\d{1,2})" dSep "(\d{1,2})" dSep "(\d{4}|\d{2})",d) {	; Jan-03-2015, 01-03-2015
		date.dd := zdigit(d2)
		date.mmm := objhasvalue(mo,d1) 
			? d1
			: mo[d1]
		date.mm := objhasvalue(mo,d1)
			? zdigit(objhasvalue(mo,d1))
			: zdigit(d1)
		date.yyyy := (d3~="\d{4}")
			? d3
			: (d3>50)
				? "19" d3
				: "20" d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"\b(\d{4})(\d{2})(\d{2})\b",d) {								; 20150103
		date.yyyy := d1
		date.mm := d2
		date.mmm := mo[d2]
		date.dd := d3
		date.date := trim(d)
	}
	
	if RegExMatch(x,"i)(\d{1,2}):(\d{2})(:\d{2})?(.*AM|PM)?",t) {						; 17:42 PM
		time.hr := zdigit(t1)
		time.min := t2
		time.sec := trim(t3," :")
		time.ampm := trim(t4)
		time.time := trim(t)
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, hr:time.hr, min:time.min, sec:time.sec, ampm:time.ampm, time:time.time}
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

ThousandsSep(x, s=",") {
; from https://autohotkey.com/board/topic/50019-add-thousands-separator/
	return RegExReplace(x, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" s)
}

WriteOut(path,node) {
	global wq
	
	filecheck()
	FileOpen(".lock", "W")															; Create lock file.
	locPath := wq.selectSingleNode(path)
	locNode := locPath.selectSingleNode(node)
	clone := locNode.cloneNode(true)													; make copy of wq.node
	
	if !IsObject(locNode) {
		eventlog("No such node <" path "/" node "> for WriteOut.")
		FileDelete, .lock															; release lock file.
		return error
	}
	
	z := new XML("worklist.xml")														; load a copy into z
	
	if !IsObject(z.selectSingleNode(path "/" node)) {									; no such node in z
		z.addElement("newnode",path)													; create a blank node
		node := "newnode"
	}
	zPath := z.selectSingleNode(path)													; find same "node" in z
	zNode := zPath.selectSingleNode(node)
	zPath.replaceChild(clone,zNode)														; replace existing zNode with node clone
	
	z.save("worklist.xml")
	wq := z
	FileDelete, .lock
	
	return
}

RemoveNode(node) {
	global wq
	q := wq.selectSingleNode(node)
	q.parentNode.removeChild(q)
	eventlog("Removed node " node)
	return
}

filecheck() {
	if FileExist(".lock") {
		err=0
		Progress, , Waiting to clear lock, File write queued...
		loop 50 {
			if (FileExist(".lock")) {
				progress, %p%
				Sleep 100
				p += 2
			} else {
				err=1
				break
			}
		}
		if !(err) {
			progress off
			return error
		}
		progress off
	} 
	return
}

;~ ~LButton::
;~ {
	;~ If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < DllCall("GetDoubleClickTime")) {
		;~ MouseGetPos, mouseXpos, mouseYpos, mouseWinID, mouseWinClass, 2			; put mouse coords into mouseXpos and mouseYpos, and associated winID
	;~ }
;~ return
;~ }

#Include CMsgBox.ahk
#Include xml.ahk
#Include sift3.ahk
