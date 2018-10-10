/*	TRRIQ - The Rhythm Recording Interpretation Query
	Disassembles HL7 and PDF files into discrete data elements
	Outputs into a format readable by HolterDB (CSV, PDF, and short PDF)
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
FileInstall, trriq.ini, trriq.ini
FileInstall, hl7.ini, hl7.ini

SplitPath, A_ScriptDir,,fileDir
user := A_UserName
IfInString, fileDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	;~ chip := httpComm("","full")
	;~ FileDelete, .\Chipotle\currlist.xml
	;~ FileAppend, % chip, .\Chipotle\currlist.xml
	isAdmin := true
	readIni("adminpaths")
	eventlog(">>>>> Started in DEVT mode.")
} else {
	FileGetTime, tmp, % A_ScriptName
	isAdmin := false
	readIni("paths")
	eventlog(">>>>> Started in PROD mode. " A_ScriptName " ver " substr(tmp,1,12))
}

/*	Read outdocs.csv for Cardiologist and Fellow names 
*/
progress,,,Scanning providers...
Docs := Object()
tmpChk := false
Loop, Read, %chipDir%outdocs.csv
{
	tmp := StrSplit(A_LoopReadLine,",","""")
	if (tmp.1="Name" or tmp.1="end" or tmp.1="") {				; header, end, or blank lines
		continue
	}
	if (tmp.4="group") {											; skip group names
		continue
	}
	if (tmp.2="" and tmp.3="" and tmp.4="") {						; Fields 2,3,4 blank = new group
		tmpGrp := tmp.1
		tmpIdx := 0
		tmpIdxG += 1
		outGrps.Insert(tmpGrp)
		continue
	}
	if !(tmp.4~="i)(seattlechildrens\.org|washington\.edu)") {		; skip non-SCH or non-UW providers
		continue
	}
	tmpIdx += 1
	tmpPrv := RegExReplace(tmp.1,"^(.*?) (.*?)$","$2, $1")			; input FIRST LAST NAME ==> LAST NAME, FIRST
	Docs[tmpGrp,tmpIdx]:=tmpPrv
	Docs[tmpGrp ".eml",tmpIdx] := tmp.4
	Docs[tmpGrp ".npi",tmpIdx] := tmp.5
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

demVals := readIni("demVals")																		; valid field names for parseClip()

indCodes := readIni("indCodes")
for key,val in indCodes
{
	tmpVal := strX(val,"",1,0,":",1)
	tmpStr := strX(val,":",1,1,"",0)
	indOpts .= tmpStr "|"
}

initHL7()

#Include HostName.ahk
progress,,,Identifying workstation...
if !(wksLoc := GetLocation()) {
	progress, off
	MsgBox, 262160, Location error, No clinic location specified!`n`nExiting
	ExitApp
}

site := getSites(wksLoc)
sites := site.tracked																	; sites we are tracking
sites0 := site.ignored																	; sites we are not tracking <tracked>N</tracked> in wkslocation
sitesLong := site.long																	; {CIS:TAB}
sitesCode := site.code																	; {"MAIN":7343} 4 digit code for sending facility
sitesFacility := site.facility															; {"MAIN":"GB-SCH-SEATTLE"}

MainLoop: ; ===================== This is the main part ====================================
{
	Loop
	{
		Gosub PhaseGUI
		WinWaitClose, TRRIQ Dashboard
		
		if (phase="Enrollment") {
			eventlog("Update Preventice enrollments.")
			CheckPreventiceWeb("Patient Enrollment")
		}
		if (phase="Inventory") {
			eventlog("Update Preventice inventory.")
			CheckPreventiceWeb("Facilities")
		}
		if (phase="Register") {
			eventlog("Start BGH registration.")
			BGHregister()
		}
		if (phase="Upload") {
			eventlog("Start Mortara preparation/upload.")
			MortaraUpload()
		}
	}
	
	ExitApp
}

PhaseGUI:
{
	phase :=
	Gui, phase:Destroy
	Gui, phase:Default
	Gui, +AlwaysOnTop

	Gui, Add, Text, x650 y15 w200
		, % "Patients registered in Preventice (" wq.selectNodes("/root/pending/enroll").length ")`n"
		.	"Last Enrollments update: " niceDate(wq.selectSingleNode("/root/pending").getAttribute("update")) "`n"
		.	"Last Inventory update: " niceDate(wq.selectSingleNode("/root/inventory").getAttribute("update")) 
	Gui, Add, GroupBox, x640 y0 w220 h65
	
	Gui, Font, Bold
	Gui, Add, Button
		, Y+10 wp h40 gPhaseTask
		, Refresh files
	Gui, Add, Button
		, Y+10 wp h40 vEnrollment gPhaseTask
		, Grab Preventice enrollments
	Gui, Add, Button
		, Y+10 wp h40 vInventory gPhaseTask
		, Grab Preventice inventory
	Gui, Add, Button
		, Y+10 wp h40 vRegister gPhaseTask
		, Register BGH EVENT RECORDER
	Gui, Add, Button
		, Y+10 wp h40 vUpload gPhaseTask
		, Prepare/Upload MORTARA HOLTER
	Gui, Font, Normal
	
	GuiControl
		, % (wksloc="Main Campus" ? "Enable" : "Disable") 
		, Grab Preventice enrollments
	GuiControl
		, % (wksloc="Main Campus" ? "Enable" : "Disable") 
		, Grab Preventice inventory
	
	Gui, Add, Tab3
		, -Wrap x10 y10 w620 h320 vWQtab +HwndWQtab
		, % (wksloc="Main Campus" ? "INBOX|" : "") "ALL|" RegExReplace(sites,"TRI\|")	; add Tab bar with tracked sites
	WQlist()
	
	;~ Menu, menuSys, Add, Scan tempfiles, scanTempFiles
	;~ Menu, menuSys, Add, Find returned devices, WQfindlost
	Menu, menuSys, Add, Change clinic location, changeLoc
	Menu, menuSys, Add, Generate late returns report, lateReport
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
	eventlog("<<<<< Session end.")
	ExitApp
}	

menuTrriq:
{
	Gui, phase:hide
	FileGetTime, tmp, % A_ScriptName
	MsgBox, 64, About..., % A_ScriptName " version " substr(tmp,1,12) "`nTerrence Chun, MD"
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

changeLoc:
{
	MsgBox, 262193, Change clinic, Current location: %wksLoc%`n`nReally change the clinic location for this PC?`n`nWill restart TRRIQ
	IfMsgBox, Ok
	{
		locationData := new xml(m_strXmlFilename)                               	  ; load xml file
		wksList := locationData.SelectSingleNode(m_strXmlWorkstationsPath)            ; retreive list of all workstations
		wksNode := wksList.selectSingleNode(m_strXmlWksNodeName "[" m_strXmlWksName "='" A_ComputerName "']")
		wksNode.parentNode.removeChild(wksNode)
		locationData.TransformXML()
		locationData.saveXML()
		eventlog("Removed wks node for " A_ComputerName)
		Reload
	}
	return
}

lateReport:
{
	str := ""
	Loop, % (ens:=wq.selectNodes("/root/pending/enroll")).length
	{
		k := ens.item(A_Index-1)
		id	:= k.getAttribute("id")
		e := readWQ(id)
		dt := A_now
		dt -= e.date, Days
		if (instr(e.dev,"BG") && (dt > 45)) || (instr(e.dev,"Mortara") && (dt > 14))  {
			str .= e.site ",""" e.prov """," e.date ",""" e.name """," e.mrn "," e.dev "`n"
		}
	}
	tmp := holterDir "late-" A_now ".csv"
	FileAppend, %str%, %tmp%
	MsgBox, 262208, Missing devices report, Report saved to:`n%tmp%
	return
}

PhaseTask:
{
	phase := A_GuiControl
	Gui, phase:Hide
	return
}

checkCitrix() {
/*	TRRIQ must be run from local machine
	local machine names begin with EWCS and Citrix machines start with PPWC
*/
	if (A_ComputerName~="EWCS") {														; running on a local machine
		return																			; return successfully
	}
	else if (A_ComputerName~="PPWC") {
		MsgBox, 4112, Environment error, TRRIQ cannot be run from Citrix/VDI`nWill now exit...
		IfMsgBox, OK
		{
			eventlog("Exiting due to Citrix environment.")
			ExitApp
		} 
	}
	else {
		eventlog("Unique machine name.")
		return
	}
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
	if (choice="xClose") {
		return
	}
	if instr(choice,"upload") {
		InputBox ,inDT,Upload log,`n`nEnter date uploaded to Preventice,,,,,,,,% niceDate(A_now)
		if (ErrorLevel) {
			return
		}
		wq := new XML("worklist.xml")
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
		wq := new XML("worklist.xml")
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
		if (reason="xClose") {
			return
		}
		if instr(reason,"Other") {
			reason := maxinput("Clear record from worklist","Enter the reason for moving this record",30)
			if (reason="") {
				return
			}
		}
		wq := new XML("worklist.xml")
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
	local k, ens, e0, id, now, dt, site, fnID, res, key, val, full, wqfiles, lvDim
	wqfiles := []
	GuiControlGet, wqDim, Pos, WQtab
	lvDim := "W" wqDimW-25 " H" wqDimH-35
	
	Progress,,,Scanning worklist...
	
	fileCheck()
	FileOpen(".lock", "W")																; Create lock file.
	wq := new XML("worklist.xml")														; refresh WQ
	loop, parse, sites0, |																; move studies from sites0 to DONE
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
	
	if (wksloc="Main Campus") {
		
	Gui, Tab, INBOX
	Gui, Add, Listview
		, % "-Multi Grid BackgroundSilver " lvDim " greadWQlv vWQlv_in hwndHLV_in"
		, filename|Name|MRN|DOB|Location|Study Date|wqid|Type|FTP
	
/*	Process each .hl7 file
*/
	loop, Files, % hl7Dir "*.hl7"
	{
		fileIn := A_LoopFileName
		x := StrSplit(fileIn,"_")
		id := findWQid(SubStr(x.5,1,8),x.3).id											; get id based on study date and mrn
		res := readWQ(id)																; wqid should always be present in hl7 downloads
		if (res.node="done") {															; skip if DONE, might be currently in process 
			eventlog("Report already done. WQlist removing " fileIn)
			FileMove, % hl7Dir "\" fileIn, .\tempfiles\%fileIn%, 1
			continue
		}
		FileGetSize,full,% hl7Dir fileIn,M
		
		LV_Add(""
			, hl7Dir fileIn																; path and filename
			, strQ(res.Name,"###", x.1 ", " x.2)										; last, first
			, strQ(res.mrn,"###",x.3)													; mrn
			, strQ(niceDate(res.dob),"###",niceDate(x.4))								; dob
			, strQ(res.site,"###","???")												; site
			, strQ(niceDate(res.date),"###",niceDate(SubStr(x.5,1,8)))					; study date
			, id																		; wqid
			, (res.dev~="BG") ? "BGH"													; extracted
			: (res.dev~="Mortara") ? "HOL"
			: "HL7"
			, (full>2)||(res.dev~="BG") ? "":"X")										; fulldisc if filesize >2 Meg
		wqfiles.push(id)
	}
	
/*	Scan Holter PDFs folder for additional files
*/
	findfullPDF()
	for key,val in pdfList
	{
		RegExMatch(val,"O)_WQ(\d+)(\w)?\.pdf",fnID)										; get filename WQID if PDF has already been renamed (fnid.1 = wqid, fnid.2 = type)
		id := fnID.1
		ftype := (fnID.2="H") ? "PDF"													; type of file based on fnID label
				: (fnID.2="Z") ? "ZIO"
				: (fnID.2="E") ? "CEM"
				: (fnID.2="M") ? "MINI"
				: "???"																	; could condense as ftype := {"H":"PDF","Z":"ZIO","E":"CEM","M":"MINI"}[fnID.2]
		if (k:=ObjHasValue(wqfiles,id)) {												; found a PDF file whose wqid matches an hl7 in wqfiles
			LV_Modify(k,"Col9","")														; clear the "X" in the FullDisc column
			continue																	; skip rest of processing
		}
		res := readwq(id)																; get values for wqid if valid, else null
		
		LV_Add(""
			, HolterDir val																; filename and path to HolterDir
			, strQ(res.Name,"###",strX(val,"",1,0,"_",1))								; name from wqid or filename
			, strQ(res.mrn,"###",strX(val,"_",1,1,"_",1))								; mrn
			, strQ(res.dob,"###")														; dob
			, strQ(res.site,"###","???")												; site
			, strQ(res.date,"###")														; study date
			, id																		; wqid
			, ftype																		; study type
			, "")																		; fulldisc present, make blank
		if (id) {
			wqfiles.push(id)															; add non-null wqid to wqfiles
		}
	}
	
	Gui, ListView, WQlv_in
		LV_ModifyCol()
		LV_ModifyCol(1,"0")																; filename and path
		LV_ModifyCol(2,"140")															; name
		LV_ModifyCol(3,"60")															; mrn
		LV_ModifyCol(4,"80")															; dob
		LV_ModifyCol(5,"60")															; site
		LV_ModifyCol(6,"80 Sort")														; date
		LV_ModifyCol(7,"2")																; wqid
		LV_ModifyCol(8,"40")															; ftype
		LV_ModifyCol(9,"40 Center")														; ftp
		
	}	; <-- finish Main Campus Inbox
	
	Gui, Tab, ALL
	Gui, Add, Listview
		, % "-Multi Grid BackgroundSilver " lvDim " gWQtask vWQlv_all hwndHLV_all"
		, ID|Enrolled|FedEx|Uploaded|MRN|Enrolled Name|Device|Provider|Site
	
	Loop, parse, sites, |
	{
		i := A_index
		site := A_LoopField
		Gui, Tab, % site
		Gui, Add, Listview
			, % "-Multi Grid BackgroundSilver " lvDim " gWQtask vWQlv"i " hwndHLV"i
			, ID|Enrolled|FedEx|Uploaded|MRN|Enrolled Name|Device|Provider
		Loop, % (ens:=wq.selectNodes("/root/pending/enroll[site='" site "']")).length
		{
			k := ens.item(A_Index-1)
			id	:= k.getAttribute("id")
			e0 := readWQ(id)
			dt := A_now
			dt -= e0.date, Days
			e0.dev := RegExReplace(e0.dev,"BodyGuardian","BG")
			if (instr(e0.dev,"BG") && (dt < 30)) {
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
			Gui, ListView, WQlv_all														
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
	
	Gui, ListView, WQlv_all
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
	global wq
	MsgBox, 4132, Find devices, Scan database for duplicate devices?`n`n(this can take a while)
	IfMsgBox, Yes
	{
		wq := new XML("worklist.xml")
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
	res.node := k.parentNode.nodeName 
	
	return res
}

readWQlv:
{
/*	Retrieve info from WQlist line
	Will be for HL7 data, or an additional file in Holter PDFs folder
*/
	agc := A_GuiControl
	if !instr(agc,"WQlv") {																; Must be in WQlv listview
		return
	}
	if !(A_GuiEvent="DoubleClick") {													; Must be double click
		return
	}
	Gui, ListView, %agc%
	if !(x := LV_GetNext()) {															; Must be on actual row
		return
	}
	LV_GetText(fileIn,x,1)																; selection filename
	LV_GetText(wqid,x,7)																; WQID
	LV_GetText(ftype,x,8)																; filetype
	SplitPath,fileIn,fnam,,,fileNam
	
	blocks := Object()																	; clear all objects
	fields := Object()
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
	obxval := Object()
	
	fldVal := readWQ(wqid)																; wqid would have been determined by parsing hl7
	fldval.wqid := wqid																	; or findFullPdf scan of extra PDFs
	
	if (ftype~="HL7|HOL|CEM|BGH") {														; hl7 file (could still be Holter or CEM)
		eventlog("===> " fnam )
		Gui, phase:Hide
		
		progress, 25 , % fnam, Extracting data
		processHL7(fnam)																; extract DDE to fldVal, and PDF into hl7Dir
		moveHL7dem()																	; prepopulate the fldval["dem-"] values
		
		progress, 50 , % fnam, Processing PDF
		gosub processHl7PDF																; process resulting PDF file
	}
	else if (ftype) {																	; Any other PDF type
		FileGetSize, fileInSize, %fileIn%
		Gui, phase:Hide
		eventlog("===> " fnam " type " ftype " (" thousandsSep(fileInSize) ").")
		gosub processPDF
	}
	else {
		Gui, phase:Hide
		eventlog("Filetype cannot be determined from WQlist (somehow).")
		
		MsgBox, 16, , Unrecognized filetype (somehow)
	}
	
	if (fldval.done) {
		epRead()																		; find out which EP is reading today
		gosub outputfiles																; generate and save output CSV, rename and move PDFs
	}
	
	gosub PhaseGUI																		; Refresh the worklist
	
	return
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
					tmpCrd := checkCrd(ptDem.Provider)
					ptDem.NPI := Docs[tmpCrd.Group ".npi",ObjHasValue(Docs[tmpCrd.Group],tmpCrd.best)]
					mdCoord.x4 := mouseXpos													; demographics grid[4,1]
					mdCoord.y1 := mouseYpos
					mdProv := true														; we have got Provider
					gosub getDemName													; extract patient name, MRN from window title 
				}																		; (this is why it must be sister or parent VM).
				if (clk.field = "Account Number") {
					ptDem["Account"] := clk.value
					eventlog("MouseGrab Account Number " clk.value ".")
					mdCoord.x1 := mouseXpos													; demographics grid[1,3]
					mdCoord.y3 := mouseYpos
					mdAcct := true														; we have got Acct Number
					gosub getDemName													; extract patient name, MRN
				}
				if (mdProv and mdAcct) {												; we have both critical coordinates
					mdX0 := 50
					mdX1 := mdWinW*0.15
					mdX2 := mdWinW*0.25
					mdX3 := mdWinW*0.35
					mdCoord.x1 := mdX0
					mdCoord.x2 := mdX0 + mdWinW*0.15
					mdCoord.x3 := mdX0 + mdWinW*0.25
					mdCoord.x4 := mdX0 + mdWinW*0.35
					mdCoord.x5 := mdX0 + mdWinW*0.45
					mdCoord.x6 := mdX0 + mdWinW*0.65
					mdCoord.x7 := mdX0 + mdWinW*0.85
					mdCoord.y2 := mdCoord.y1+(mdCoord.y3-mdCoord.y1)/2									; determine remaning row coordinate
					
					Gui, fetch:hide														; grab remaining demographic values
					BlockInput, On														; Prevent extraneous input
					ptDem["MRN"] := mouseGrab(mdCoord.x1,mdCoord.y2).value
					ptDem["DOB"] := mouseGrab(mdCoord.x2,mdCoord.y2).value
					ptDem["Sex"] := mouseGrab(mdCoord.x4,mdCoord.y1).value
					eventlog("MouseGrab other fields. MRN=" ptDem["MRN"] " DOB=" ptDem["DOB"] " Sex=" ptDem["Sex"] ".")
					
					tmp := mouseGrab(mdCoord.x6,mdCoord.y3)										; grab Encounter Type field
					ptDem["Type"] := tmp.value
					if (ptDem["Type"]="Outpatient") {
						ptDem["Loc"] := mouseGrab(mdCoord.x7-mdX0-30,mdCoord.y2).value				; most outpatient locations are short strings, click the right half of cell to grab location name
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
		val1 := RegExReplace(val1,"Legal Sex|Birth Sex","Sex")
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
	EncNum := ptDem["Account"]						; we need these non-array variables for the Gui statements
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
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("Account","\d{8,}",1), Encounter #
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
	
	if (instr(ptDem.Provider," ") && !instr(ptDem.Provider,",")) {						; somehow string passed in wrong order
		tmp := parseName(ptDem.Provider)
		ptDem.Provider := tmp.Last ", " tmp.First
	}
	matchProv := checkCrd(ptDem.Provider)
	if !(ptDem.Provider) {																; no provider? ask!
		gosub getMD
		eventlog("New provider field " ptDem.Provider ".")
	} else if (matchProv.fuzz > 0.10) {													; Provider not recognized
		eventlog(ptDem.Provider " not recognized (" matchProv.fuzz ").")
		if (ptDem.Type~="i)(Inpatient|Observation|Emergency|Day Surg)") {
			ptDem.EncDate := substr(A_now,1,8)											; Set date to today
			gosub assignMD																; Inpt, ER, DaySurg, we must find who recommended it from the Chipotle schedule
			eventlog(ptDem.Type " location. Provider assigned to " ptDem.Provider ".")
		} else {
			gosub getMD																	; Otherwise, ask for it.
			eventlog("Provider set to " ptDem.Provider ".")
		}
	} else {																			; Provider recognized
		eventlog(ptDem.Provider " matches " matchProv.Best " (" (1-matchProv.fuzz)*100 ").")
		ptDem.Provider := matchProv.Best
	}
	tmpCrd := checkCrd(ptDem.provider)													; Make sure we have most current provider
	ptDem.NPI := Docs[tmpCrd.Group ".npi",ObjHasValue(Docs[tmpCrd.Group],tmpCrd.best)]
	ptDem["Account"] := EncNum															; make sure array has submitted EncNum value
	FormatTime, EncDt, %EncDt%, MM/dd/yyyy												; and the properly formatted date 06/15/2016
	ptDem.EncDate := EncDt
	ptDemChk := (ptDem["nameF"]~="i)[A-Z\-]+") && (ptDem["nameL"]~="i)[A-Z\-]+") 		; valid names
			&& (ptDem["mrn"]~="\d{6,7}") && (ptDem["Account"]~="\d{8,}") 				; valid MRN and Acct numbers
			&& (ptDem["DOB"]~="[0-9]{1,2}/[0-9]{1,2}/[1-2][0-9]{3}") && (ptDem["Sex"]~="^[MF]") 		; valid DOB and Sex
			&& (ptDem["Loc"]) && (ptDem["Type"])										; Loc and type is not null
			&& (ptDem["Provider"]~="i)[a-z]+") && (ptDem["EncDate"])					; prov any string, encDate not null
	if !(ptDemChk) {																	; all data elements must be present, otherwise retry
		eventlog("Data incomplete."
			. ((ptDem["nameF"]) ? "" : " nameF")
			. ((ptDem["nameL"]) ? "" : " nameL")
			. ((ptDem["mrn"]) ? "" : " MRN")
			. ((ptDem["Account"]) ? "" : " EncNum")
			. ((ptDem["DOB"]) ? "" : " DOB")
			. ((ptDem["Sex"]) ? "" : " Sex")
			. ((ptDem["Loc"]) ? "" : " Loc")
			. ((ptDem["Type"]) ? "" : " Type")
			. ((ptDem["EncDate"]) ? "" : " EncDate")
			. ((ptDem["Provider"]) ? "" : " Provider")
			. ".")
		MsgBox,, % "Data incomplete. Try again", % ""
			. "First name " ptDem["nameF"] "`n"
			. "Last name " ptDem["nameL"] "`n"
			. "MRN " ptDem["mrn"] "`n"
			. "Account number " ptDem["Account"] "`n"
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
	Gui, ind:Destroy
	Gui, ind:+AlwaysOnTop
	Gui, ind:font, s12
	Gui, ind:Add, Text, , % "Enter indications: " ptDem["Indication"]
	Gui, ind:Add, ListBox, w360 r12 vIndChoices 8, %indOpts%
	Gui, ind:Add, Button, gindSubmit, Submit
	Gui, ind:Show, Autosize, Enter indications
	return
}

indClose:
{
	Gui, ind:Destroy
	fetchQuit := true
	return
}

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

CheckPreventiceWeb(win) {
	global phase
	checkCitrix()
	
	str := {}
	str.Enrollment := {dlg:"Enrollment / Submitted Patients"
		, url:"https://secure.preventice.com/Enrollments/EnrollPatients.aspx?step=2"
		, tbl:"ctl00_mainContent_PatientListSubmittedCtrl1_RadGridPatients_ctl00"
		, changed:"ctl00_mainContent_PatientListSubmittedCtrl1_lblTotalCountMessage"
		, btn:"ctl00_mainContent_PatientListSubmittedCtrl1_btnNextPage"
		, fx:"ParsePreventiceEnrollment"}
	str.Inventory := {dlg:"Facility`nInventory Status`nDevice in Hand (Enrollment not linked)"
		, url:"https://secure.preventice.com/Facilities/"
		, tbl:"ctl00_mainContent_InventoryStatus_userControl_gvInventoryStatus_ctl00"
		, changed:"ctl00_mainContent_InventoryStatus_userControl_gvInventoryStatus_ctl00_Pager"
		, btn:"rgPageNext"
		, fx:"ParsePreventiceInventory"}
	
	while !(WinExist(win))																; expected IE window title not present
	{
		MsgBox,4161,Update Preventice %phase%
			, % "Navigate on Preventice website to:`n`n"
			.	str[phase].dlg "`n`n"
			.	"Click OK when ready to proceed"
		IfMsgBox, Cancel
		{
			return
		}
	}
	
	prvFunc := str[phase].fx
	wb := IEGet(win)
	
	loop
	{
		tbl := wb.document.getElementById(str[phase].tbl)
		if !IsObject(tbl) {
			progress, off
			MsgBox No match
			return
		}
		progress,,,Scanning page %A_index% ...
		
		tbl := tbl.getElementsByTagName("tbody")[0]
		clip := tbl.innertext
		if (clip=clip0) {
			progress, off
			MsgBox,4144,, Reached the end of novel records.`n`n%phase% update complete!
			break
		}
		
		done := %prvFunc%(tbl)		; parsePreventiceEnrollment() or parsePreventiceInventory()
		if (done=0) {
			progress, off
			MsgBox,4144,, Reached the end of novel records.`n`n%phase% update complete!
			break
		}
		clip0 := clip																	; set the check for repeat copy
		
		PreventiceWebPager(wb,str[phase].changed,str[phase].btn)
	}
	
	wb.navigate(str[phase].url)															; refresh first page
	ComObjConnect(wb)																	; release wb object
	return
}

PreventiceWebPager(wb,chgStr,btnStr) {
	global phase
	
	if (phase="Enrollment") {
		wb.document.getElementById(btnStr).click() 										; click when id=btnStr
	}
	if (phase="Inventory") {
		wb.document.getElementsByClassName(btnStr)[0].click() 							; click when class=btnstr
	}
	pg0 := wb.document.getElementById(chgStr).innerText
	
	loop, 100																			; wait up to 100*0.05 = 5 sec
	{
		pg := wb.document.getElementById(chgStr).innerText
		progress,% A_index
		if (pg != pg0) {
			break
		}
		sleep 50
	}

	return
}

parsePreventiceEnrollment(tbl) {
	global wq
	
	lbl := ["name","mrn","date","dev","prov"]
	done := 0
	fileCheck()
	wq := new XML("worklist.xml")														; refresh WQ
	FileOpen(".lock", "W")																; Create lock file.
	
	loop % (trows := tbl.getElementsByTagName("tr")).length								; loop through rows
	{
		r_idx := A_index-1
		trow := trows[r_idx]
		tcols := trow.getElementsByTagName("td")
		res := []
		loop % lbl.length()																; loop through cols
		{
			c_idx := A_Index-1
			res[lbl[A_index]] := trim(tcols[c_idx].innertext)
		}
		tmp := parseDate(res.date)
		date := tmp.YYYY tmp.MM tmp.DD
		count ++
		
		if IsObject(wq.selectSingleNode("/root/pending/enroll[mrn='" res.mrn "'][dev='" res.dev "']")) {	; S/N is currently in use
			eventlog("Enrollment for " res.mrn " " res.name " " date " already exists in Pending.")
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
			if ((e0.fuzzName>85)||(e0.fuzzMRN>90)) {									; close match for either NAME or MRN
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
		newID := "/root/pending/enroll[@id='" id "']"
		wq.addElement("enroll","/root/pending",{id:id})
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
	
	return done																			; returns number of matches, or 0 (error) if no matches
}

parsePreventiceInventory(tbl) {
/*	Parse Preventice website for device inventory
	Add unique ser nums to /root/inventory/dev[@ser]
	These will be removed when registered
*/
	global wq
	
	lbl := ["button","model","ser"]
	wq := new XML("worklist.xml")														; refresh WQ
	
	wqtime := wq.selectSingleNode("/root/inventory").getAttribute("update")
	if !(wqTime) {
		wq.addElement("inventory","/root")
		eventlog("Created new Inventory node.")
	}
	
	loop % (trows := tbl.getElementsByTagName("tr")).length								; loop through rows
	{
		r_idx := A_index-1
		trow := trows[r_idx]
		tcols := trow.getElementsByTagName("td")
		res := []
		loop % lbl.length()																; loop through cols
		{
			c_idx := A_Index-1
			res[lbl[A_index]] := trim(tcols[c_idx].innertext)
		}
		if IsObject(wq.selectSingleNode("/root/inventory/dev[@ser='" res.ser "']")) {	; already exists in Inventory
			continue
		}
		wq.addElement("dev","/root/inventory",{model:res.model,ser:res.ser})
		eventlog("Added new Inventory dev " res.ser)
	}

	loop, % (devs := wq.selectNodes("/root/inventory/dev")).length						; Find dev that already exist in Pending
	{
		k := devs.item(A_Index-1)
		ser := k.getAttribute("ser")
		if IsObject(wq.selectSingleNode("/root/pending/enroll[dev='BodyGuardian Heart - " ser "']")) {	; exists in Pending
			k.parentNode.removeChild(k)
			eventlog("Removed inventory ser " ser)
		}
	}
	
	wq.selectSingleNode("/root/inventory").setAttribute("update",A_now)					; set pending[@update] attr
	
	writeout("/root","inventory")
	
	return true
}

findWQid(DT:="",MRN:="",ser:="",name:="") {
/*	DT = 20170803
	MRN = 123456789
	ser = BodyGuardian Heart - BG12345, or Mortara H3+ - 12345
	name = LAST, FIRST
*/
	global wq
	
	if IsObject(x := wq.selectSingleNode("//enroll[date='" DT "'][mrn='" MRN "']")) {		; Perfect match DT and MRN
	} else if IsObject(x := wq.selectSingleNode("//enroll"									; or matches S/N and MRN
		. "[dev='Mortara H3+ - " DT "'][mrn='" MRN "']")) {
	} else if IsObject(x := wq.selectSingleNode("//enroll[mrn='" MRN "']")) {				; or matches MRN only 
	} else if (name) && IsObject(x := wq.selectSingleNode("//enroll[name='" name "']")) {	; or neither, find matching name 
	}

	return {id:x.getAttribute("id"),node:x.parentNode.nodeName}								; will return null (error) if no match
}

scanTempfiles() {
	global wq
	count := 0
	
	filecheck()
	wq := new XML("worklist.xml")
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
		wq.addElement("scantemp",newID,A_Now)
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
	checkCitrix()
	
	if !WinExist("ahk_exe WebUploadApplication.exe") {									; launch Mortara Upload app from site if not running
		wb := ComObjCreate("InternetExplorer.Application")								; webbrowser object
		wb.Navigate("https://h3.preventice.com/WebUploadApplication.application")		; open direct link to WebUploadApplication.application
		ComObjConnect(wb)																; disconnect the webbrowser object
		WinWait, Mortara Web Upload, , 5												; wait 5 seconds for window to appear
	}
	
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
		
		wuDir := {}
		Loop, files, % WebUploadDir "Data\*", D											; Get the most recently created Data\xxx folder
		{
			loopDate := A_LoopFileTimeModified
			loopName := A_LoopFileLongPath
			if (loopDate>=wuDir.Date) {
				wuDir.Date := loopDate
				wuDir.Full := loopName
			}
		}
		wuDir.Short := strX(wuDir.Full,"\",0,1,"",0)
		eventlog("Found WebUploadDir " wuDir.Short )
		FileReadLine, wuRecord, % wuDir.Full "\RECORD.LOG", 1
		FileReadLine, wuDevice, % wuDir.Full "\DEVICE.LOG", 1
		wuDir.MRN := trim(RegExReplace(wuRecord,"i)Patient ID:"))
		wuDir.Ser := substr(wuDevice,-4)
		eventlog("Data files: wuDirSer " wuDir.Ser ", MRN " wuDir.MRN)
		if !(serNum=wuDir.Ser) {
			eventlog("Serial number mismatch.")
			MsgBox, 262160, Device error, Device mismatch!`n`nTry again.
			return
		}
		
		wq := new XML("worklist.xml")													; refresh WQ
		wqStr := "/root/pending/enroll[dev='Mortara H3+ - " SerNum "'][mrn='" wuDir.MRN "']"
		wqTR:=wq.selectSingleNode(wqStr)
		if IsObject(wqTR.selectSingleNode("acct")) {									; S/N exists, and valid
			pt := readwq(wqTR.getAttribute("id"))
			ptDem["mrn"] := pt.mrn														; fill ptDem[] with values
			ptDem["loc"] := pt.site
			ptDem["date"] := pt.date
			ptDem["Account"] := RegExMatch(pt.acct,"([[:alpha:]]+)(\d{8,})",z) ? z2 : pt.acct
			ptDem["nameL"] := strX(pt.name,"",0,1,",",1,1)
			ptDem["nameF"] := strX(pt.name,",",1,1,"",0)
			ptDem["Sex"] := pt.sex
			ptDem["dob"] := pt.dob
			ptDem["Provider"] := pt.prov
			ptDem["Indication"] := pt.ind
			ptDem["loc"] := z1
			ptDem["wqid"] := wqTR.getAttribute("id")
			eventlog("Found valid registration for " pt.name " " pt.mrn " " pt.date)
			MsgBox, 262193
				, Match!
				, % "Found valid registration for:`n" pt.name "`n" pt.mrn "`n" pt.date
			IfMsgBox, Cancel
			{
				eventlog("Cancelled GUI.")
				return
			}
		} else {																		; no valid S/N exists
			gosub getDem																; fill ptDem[] with values
			if (fetchQuit=true) {
				fetchQuit:=false
				eventlog("Cancelled getDem.")
				return
			}
			ptDem["muphase"] := "upload"
			muWqSave(SerNum)
			eventlog(ptDem["muphase"] ": " sernum " registered to " ptDem["mrn"] " " ptDem["nameL"] ".") 
			wqStr := "/root/pending/enroll[dev='Mortara H3+ - " SerNum "'][mrn='" ptDem["mrn"] "']"
		}
		MorUIfill(mu_UI.TRct,muWinID)
		
		Gui, muTm:Add, Progress, w150 h6 -smooth hwndMtCt 0x8
		Gui, muTm:+ToolWindow
		Gui, muTm:Show, AutoSize, Close to cancel upload...
		SetTimer, muTimer, 50
		ptDem.timer := false
		
		loop
		{
			if FileExist(wuDir.Full "\Uploaded.txt") {
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
		MsgBox, 262208, Transfer, Successful data upload to Preventice.
	}
	
	if (Tabnum=2) {																		; PREPARE MEDIA TAB
		eventlog("Prepare media selected.")
		mu_UI := MorUIgrab()
		
		gosub getDem
		if (fetchQuit=true) {
			fetchQuit:=false
			eventlog("Cancelled getDem.")
			return
		}
		getPatInfo()																	; grab remaining demographics for Preventice registration
		if (fetchQuit=true) {
			eventlog("Cancelled getPatInfo.")
			return
		}
		InputBox, note, Fedex, `n`n`n`n Enter FedEx return sticker number
		if (RegExMatch(note,"((\d\s*){12})",fedex)) {
			fedex := RegExReplace(fedex1," ")
			ptDem["fedex"] := fedex
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
		
		wq := new XML("worklist.xml")													; refresh WQ
		ptDem["muphase"] := "prepare"
		ptDem["hookup"] := "Office"
		muWqSave(SerNum)
		eventlog(ptDem["muphase"] ": " sernum " registered to " ptDem["mrn"] " " ptDem["nameL"] ".") 
		
		registerPreventice()
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
	
	if (ptDem.EncDate) {
		tmp := parsedate(ptDem.EncDate)
		ptDem.date := tmp.YYYY tmp.MM tmp.DD
	}
	
	id := A_TickCount
	ptDem["model"] := "Mortara H3+"
	ptDem["ser"] := sernum
	ptDem["dev"] := ptDem.model " - " sernum
	ptDem["wqid"] := id
	
	wq.addElement("enroll","/root/pending",{id:id})
	newID := "/root/pending/enroll[@id='" id "']"
	wq.addElement("date",newID,(ptDem["date"]) ? ptDem["date"] : substr(A_now,1,8))
	wq.addElement("name",newID,ptDem["nameL"] ", " ptDem["nameF"])
	wq.addElement("mrn",newID,ptDem["mrn"])
	wq.addElement("sex",newID,ptDem["Sex"])
	wq.addElement("dob",newID,ptDem["dob"])
	wq.addElement("dev",newID,ptDem["dev"])
	wq.addElement("prov",newID,ptDem["Provider"])
	wq.addElement("site",newID,sitesLong[ptDem["loc"]])										; need to transform site abbrevs
	wq.addElement("acct",newID,ptDem["loc"] ptDem["Account"])
	wq.addElement("ind",newID,ptDem["Indication"])
	if (ptDem.fedex) {
		wq.addElement("fedex",newID,ptDem["fedex"])
	}
	wq.addElement(ptDem["muphase"],newID,{user:A_UserName},A_now)
	
	filedelete, .lock
	writeOut("/root/pending","enroll[@id='" id "']")
	
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

registerPreventice() {
	global wq, ptDem, fetchQuit, hl7out, hl7OutDir, indCodes, sitesCode, sitesFacility
	
	hl7time := A_Now
	hl7out := Object()
	buildHL7("MSH"
		,"^~\&"
		,"TRRIQ"
		,sitesCode
		,sitesFacility
		,"PREVENTICE"
		,hl7time
		,"TECH"
		,"ORM^O01"
		,ptDem["wqid"]
		,"T"
		,"2.3")
	
	tmpDOB := parseDate(ptDem.dob)
	buildHL7("PID"
		, ptDem.MRN
		, ptDem.MRN
		, ""
		, ptDem.nameL "^" ptDem.nameF . strQ(ptDem.nameMI,"^###")
		, ""
		, tmpDOB.yyyy . tmpDOB.mm . tmpDOB.dd
		, substr(ptDem.sex,1,1)
		, ""
		, ""
		, ptDem.Addr1 "^" ptDem.Addr2 "^" ptDem.city "^" ptDem.state "^" ptDem.zip
		, ""
		, ptDem.phone
		, ""
		, ""
		, ""
		, ""
		, ptDem.account
		, "")
	
	tmpPrv := parseName(ptDem.provider)
	buildHL7("PV1"
		, ptDem.type
		, ptDem.loc
		, ""
		, ""
		, ""
		, ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, ""
		, ""
		, ""
		, ""
		, ""
		, ""
		, ""
		, ""
		, ""
		, ""
		, ptDem.account)
	
	buildHL7("IN1",
		, "N/A"
		, "" ;"Insurance Company ID"
		, "Seattle Childrens - GB" ;"Insurance Company Name"
		, "" ;"Insurance Company Address"
		, "" ;"Insurance Co Contact Person"
		, "" ;"Insurance Co Phone Number"
		, "" ;"Group Number"
		, "" ;"Group Name"
		, "" ;"Insureds Group Emp ID"
		, "" ;"Insureds Group Emp Name"
		, "" ;"Plan Effective Date"
		, "" ;"Plan Expiration Date"
		, "" ;"Authorization Information"
		, "" ;"Plan Type"
		, ptDem.nameL "^" ptDem.nameF . strQ(ptDem.nameMI,"^###")
		, "Self"
		, tmpDOB.yyyy . tmpDOB.mm . tmpDOB.dd
		, "" ;ptDem.Addr1 "^" ptDem.Addr2 "^" ptDem.city "^" ptDem.state "^" ptDem.zip
		, "" ;"Assignment of Benefits"
		, "" ;"Coordination of Benefits"
		, "" ;"Primary Payor"
		, "" ;"Notice of Admission Code"
		, "" ;"Notice of Admission Date"
		, "" ;"Report of Eligibility Flag"
		, "" ;"Report of Eligibility Date"
		, "" ;"Release Information Code"
		, "" ;"Pre-Admit Cert (PAC)"
		, "" ;"Verification Date/Time"
		, "" ;"Verification By"
		, "" ;"Type of Agreement Code"
		, "" ;"Billing Status"
		, "" ;"Lifetime Reserve Days"
		, "" ;"Delay Before L R Day"
		, "" ;"Company Plan Code"
		, "" ;"Policy Number"
		, "" ;"Bill Type"
		, "" ;"Blank"
		, "" ;"Blank"
		, "" ;"Blank"
		, "" ;"Blank"
		, "" ;"Blank"
		, "" ;"Blank"
		, "" ;"Blank"
		, "" ;"Blank"
		, "" ;"Blank"
		, "")
	
	buildHL7("ORC","")
	
	buildHL7("OBR"
		, ptDem.account
		, ""
		, strQ((ptDem.model~="Mortara") ? 1 : "","Holter^Holter")
		. strQ((ptDem.model~="BodyGuardian") ? 1 : "","CEM^CEM")
		, ""
		, ""
		, hl7time
		, ""
		, ""
		, ""
		, "ANCILLARY"
		, ""
		, ""
		, ""
		, ""
		, ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, "206-987-2015"
		, "","","","","","","","","","","")
	
	tmpInd := ptDem.indication
	loop, parse, tmpInd, |
	{
		indIdx := ""
		indSeg := A_LoopField
		for key,val in indCodes
		{
			indVal := strX(val,"",1,0,":",1)
			indStr := strX(val,":",1,1,"",0)
			if (indSeg=indStr) {
				indIdx := indVal
				break
			}
		}
		
		buildHL7("DG1"
			, ""
			, indIdx
			, indSeg)
	}
	
	buildHL7("OBX"
		, "ST", "12915^Service Type", ""
		, strQ((ptDem.model~="Mortara") ? 1 : "","Holter")
		. strQ((ptDem.model~="BodyGuardian") ? 1 : "","CEM") )
	
	buildHL7("OBX"
		, "ST", "12916^Device", "", ptDem.model)
	
	buildHL7("OBX"
		, "ST", "12919^Serial Number", "", ptDem.ser)
	
	buildHL7("OBX"
		, "ST", "12917^Hookup Location", "", ptDem.Hookup)
	
	buildHL7("OBX"
		, "ST", "12918^Deploy Duration (In Days)", ""
		, (ptDem.model~="Mortara" ? "1" : "")
		. (ptDem.model~="BodyGuardian" ? "30" : "") )
	
	fileNm := ptDem.nameL "_" ptDem.nameF "_" ptDem.mrn "-" hl7time ".txt"
	FileAppend, % hl7Out.msg, % ".\tempfiles\" fileNm
	FileCopy, % ".\tempfiles\" fileNm , % hl7OutDir . fileNm
	eventlog("Preventice registration completed: " fileNm)
	MsgBox, 262208, Preventice registration, Successful device registration!
	return
}

BGHregister() {
	global wq, ptDem, fetchQuit
	checkCitrix()
	
	MsgBox, 262177, Event recorder, Start BGH event recorder registration?
	IfMsgBox, Cancel
	{
		return
	}
	
	ptDem := object()																	; need to initialize ptDem
	fetchQuit := false
	gosub getDem																		; need to grab CIS demographics
	if (fetchQuit=true) {
		eventlog("Cancelled getDem.")
		return
	}
	getPatInfo()																		; grab remaining demographics for Preventice registration
	if (fetchQuit=true) {
		eventlog("Cancelled getPatInfo.")
		return
	}
	
	i := cMsgBox("Hook-up","Delivery type","Office|Home")
	if (i="Home") {
		ptDem["hookup"] := "Home"
		ptDem["model"] := "BodyGuardian Heart"
		eventlog("BGH home registration for " ptDem["mrn"] " " ptDem["nameL"] ".") 
	} else {																			; either Office or [X]
		ptDem["hookup"] := "Office"
		ptDem.ser := selectDev()														; need to grab a BGH ser num
		if (ptDem.ser="") {
			eventlog("Cancelled selectDev.")
			return
		}
		ptDem.model := wq.selectSingleNode("/root/inventory/dev[@ser='" ptDem.ser "']").getAttribute("model")
		if !(ptDem.model) {
			i := cMsgBox("Recorder type","Which recorder?","BodyGuardian Heart")
			if (i="xClose") {
				return
			} else {
				ptDem.model := i
			}
		}
		removeNode("/root/inventory/dev[@ser='" ptDem.ser "']")							; take out of inventory
		writeOut("/root","inventory")
		eventlog(ptDem.ser " registered to " ptDem["mrn"] " " ptDem["nameL"] ".") 
	}
	bghWqSave(ptDem.ser)																; write to worklist.xml
	
	registerPreventice()
}

selectDev() {
/*	User starts typing any number from label
	and ComboBox offers available devices
*/
	global wq, selBox, selEdit, selBut, fetchQuit
	static typed, devs, ser
	typed := devs := ser :=
	
	loop, % (k:=wq.selectNodes("/root/inventory/dev")).length							; Add all ser nums to devs string
	{
		i := k.item(A_Index-1).getAttribute("ser")
		if !(i) {
			continue
		}
		devs .= i "|"																	; generate listbox menu
	}
	devs := trim(devs," |`r`n")

	Gui, dev:Destroy
	Gui, dev:Default
	Gui, -MinimizeBox
	Gui, Add, Text, w180 +Wrap
		, % "Type some digits from the device serial number "
		. "until there is only one item, or type the full serial number"
	Gui, Font, s12
	Gui, Add, Edit, vselEdit gSelDevCount
	Gui, Add, ListBox, h100 vSelBox -vScroll Disabled, % devs							; listbox and button
	Gui, Add, Button, h30 vSelBut gSelDevSubmit Disabled, Submit						; disabled by default
	Gui, Show, AutoSize, Select device
	Gui, +AlwaysOnTop
	
	winwaitclose, Select device
	Gui, dev:Destroy
	
	return choice
	
	selDevCount:
	{
		GuiControlGet, typed, , selEdit													; get selEdit contents on each char
		tmpDev := ""
		ct := 0
		tmp := []
		tmp := StrSplit(devs,"|")														; split all devs into array
		loop, % tmp.count()
		{
			i := tmp[A_index]
			if instr(i,typed) {															; item contains typed string
				tmpDev .= "|" i 														; add to tmpdev menu
				ct ++																	; increment counter
			}
		}
		tmpDev:=tmpDev ? tmpDev : "|"
		GuiControl, , selBox, % tmpDev													; update selBox menu
		
		if (ct=1) {																		; only one remaining match
			GuiControl, Enable, SelBut													; activate Submit button
			GuiControl, Enable, SelBox
			GuiControl, Choose, selBox, 1												; highlight remaining match
			
		} else if (typed~="i)^(BG)?\d{7}$") {											; typed full ser num
			GuiControl, Enable, SelBut													; activate button
			
		} else {																		; otherwise
			GuiControl, Disable, SelBut													; disable button
			GuiControl, Disable, SelBox													; and listbox
		}
		return
	}
	
	selDevSubmit:
	{
		GuiControlGet, boxed, , selBox													; get values from box and edit
		GuiControlGet, typed, , selEdit
		choice := (boxed) ? boxed : "BG" RegExReplace(typed,"[[:alpha:]]")
		if !(choice~="^BG\d{7}$") {														; ignore if doesn't match full ser num
			return
		}
		Gui, dev:Destroy
		return
	}

}

getPatInfo() {
/*	Identify Patient Info page in CIS
	Get window dimensions, activate window H50% X80%, copy to clipboard
	Parse address block
*/
	global wq, ptDem, fetchQuit
	
;	Make sure a CIS patient window exists
	Loop
	{
		if (winID := WinExist("Opened by")) {											; break out if so
			break
		}
		MsgBox, 4149, Window error, Must open CIS to proper patient
		IfMsgBox, Retry
		{
			continue																	; try again
		} else {
			fetchQuit := true
			return
		}
	}
	MouseGetPos,mouseX,mouseY															; get original mouse coords
	WinActivate, ahk_id %winID%															; activate the CIS patient window
	WinGetActiveStats, winTitle, winW, winH, winX, winY									; and get dimensions
	
;	Make sure we are on Patient Summary / Contacts, either inpatient or outpatient
	Loop
	{
		WinActivate, ahk_id %winID%
		clipboard := 
		MouseClick, Left, % winW-40, % 0.6*winH											; click just inside window
		
		Send, ^a^c																		; Select All, Copy
		sleep 200																		; need to pause to fill clipboard
		txt := Clipboard
		MouseClick, Left, % winW-40, % 0.6*winH+10										; click again to deselect all
		MouseMove, mouseX, mouseY														; move back to original coords
		if instr(txt,"Patient contact info") {
			break																		; break out of this loop
		}
		MsgBox, 4149, Window error, Navigate to:`n   * Patient Summary / Contacts
		IfMsgBox, Retry
		{
			continue
		} else {
			fetchQuit := true
			return
		}
	}
	
	ptInfo := cleanBlank(stregX(txt,"i)Patient contact info.*?\R+",1,1,"i)Family contact info",1))
	nameLine := strX(ptInfo,"",1,0,"`n",1)
	prefName := trim(stregX(nameLine,"i)Pref.*? name:",1,1,"\R+",1))
	if !instr(nameLine, ptDem.Name) {													; fetched ptInfo must contain ptDem.name
		fetchQuit := true
		eventlog("Fetched info does not match " ptDem.Name)
		return
	}
	homePhoneLine := stregX(ptInfo,"i)Home Phone:",1,1,"\R+",1)
	RegExMatch(homePhoneLine,"O)(\d{3})[^\d]+(\d{3})[^\d]+(\d{4})",ph)
	ptDem.phone := ph.value(1) "-" ph.value(2) "-" ph.value(3)

	famInfo := cleanBlank(stregX(txt "<<<<<","i)Family contact info.*?\R+",1,1,"<<<<<",1))
	relStr := "Father|Mother|Grand|Aunt|Uncle|Foster|Parent|Sibling|Cousin|Relative|Step|Adult"
	rel := Object()
	loop, parse, famInfo, `n,`r
	{
		i := A_LoopField
		if (i~="\(" relStr) {															; line contains "(Mother"
			ct ++																		; increment counter
			rel[ct] := object()															; create a rel index object
			rel[ct].name := strX(i,"",1,1,"(",1,1)										; get name string
			continue
		}
		if (i~="Home:") {
			RegExMatch(i,"O)(\d{3})[^\d]+(\d{3})[^\d]+(\d{4})",ph)
			rel[ct].phone := ph.value(1) "-" ph.value(2) "-" ph.value(3)
			continue
		}
		if !(i~="i)("
			. "Legal guardian|"															; skip lines containing these strings
			. "Birth certificate|"
			. "Comment|"
			. "Lives with|"
			. "Custody|"
			. "Mobile:|"
			. "Work:|"
			. "Inpatient|"
			. "Emergency|"
			. "Adoption|"
			. "^\s*$|"
			. "^>)")
		{
			rel[ct].addr .= i "`n"														; add address lines to each relative index string
		}
	}
	loop, % rel.MaxIndex()
	{
		i := A_index
		loop, % rel.MaxIndex()															; compare against all other addresses
		{
			j := A_Index
			if (i=j) {																	; do not compare to self
				continue
			}
			if (rel[j].phone != ptDem.phone) {
				rel.delete(j)															; remove if doesn't match patient's home phone number
			}
			if (rel[i].addr = rel[j].addr) {
				rel.delete(j)															; remove duplicate addresses
			}
		}
		if (rel[i].addr = "") {
			rel.Delete(i)																; remove entries with no address
		}
	}
	loop, % rel.MaxIndex()
	{
		nm .= A_index ") " rel[A_index].name "|"										; generate parent name menu for cmsgbox
	}
	if (rel.MaxIndex() > 1) {
		eventlog("Multiple potential parent matches (" rel.MaxIndex() ").")
		q := cmsgbox("Parent","Who is the guarantor?",trim(nm,"|"))
		if (q="xClose") {
			fetchQuit:=true
			return
		}
		choice := strX(q,"",1,1,")",1,1)
	} else {
		choice := 1
	}
	
	ptDem.parent := rel[choice].Name
	ptDem.parentL := parseName(ptDem.parent).last
	ptDem.parentF := parseName(ptDem.parent).first
	
	addrLine := 0
	loop, parse, % rel[choice].addr, `n,`r												; parse selected addr string
	{
		i := cleanspace(A_LoopField)
		if (i~=", [A-Z]{2} \d{5}") {													; matches City, State Zip
			ptDem.city := trim(stregX(i,"",1,0,", ",1))
			ptDem.state := trim(stregX(i,", ",1,1," ",1))
			ptDem.zip := trim(stregX(i "<<<",", [A-Z]{2} ",1,1,"<<<",1))
			continue
		} 
		else 																			; everything else is an addr string
		{
			addrLine ++
			addr := "addr" addrLine
			ptDem[addr] := trim(i)
		}
	}
	if (ptDem.addr1~="i)P[\. ]+O[\. ]+Box") {
		InputBox, addr1, Cannot use P.O. Box,`n`nEnter valid street address
		InputBox, addr2, Cannot use P.O. Box,`n`nEnter city,,,,,,,,% ptDem.city
		if (addr1) {
			ptDem.addr1 := addr1
			eventlog("Replaced PO box with valid address.")
		} else {
			fetchQuit := true
			return
		}
	}
	
	MsgBox, 4164
		, Patient contact information
		, % "Retrieved info `n`n"
		. "Patient name: " ptDem.nameL ", " ptDem.nameF "`n"
		. "Patient MRN: " ptDem.mrn "`n"
		. "Patient DOB: " ptDem.DOB "`n"
		. "Parent: " ptDem.parentL ", " ptDem.parentF "`n"
		. "Address:`n"
		. "   " ptDem.addr1 "`n"
		. strQ(ptDem.addr2,"   ###`n")
		. "   " ptDem.city ", " ptDem.state " " ptDem.zip "`n"
		. "Phone: " ptDem.phone "`n`n"
		. "Provider: " ptDem.provider "`n"
		. "Encounter date: " ptDem.encDate "`n"
		. "Site: " ptDem.loc
	IfMsgBox, Yes
	{
		eventlog("Accepted patient address info.")
		fetchQuit := false
	} else {
		fetchQuit := true
	}
	return
}

bghWqSave(sernum) {
	global wq, ptDem, user, sitesLong
	
	id := A_TickCount 
	tmpDT := parsedate(ptDem.EncDate)
	ptDem["date"] := tmpDT.YYYY tmpDT.MM tmpDT.DD
	ptDem["dev"] := ptDem.model " - " ptDem.ser
	ptDem["wqid"] := id
	
	wq.addElement("enroll","/root/pending",{id:id})
	newID := "/root/pending/enroll[@id='" id "']"
	wq.addElement("date",newID,(ptDem["date"]) ? ptDem["date"] : substr(A_now,1,8))
	wq.addElement("name",newID,ptDem["nameL"] ", " ptDem["nameF"])
	wq.addElement("mrn",newID,ptDem["mrn"])
	wq.addElement("sex",newID,ptDem["Sex"])
	wq.addElement("dob",newID,ptDem["dob"])
	wq.addElement("dev",newID,ptDem["dev"])
	wq.addElement("prov",newID,ptDem["Provider"])
	wq.addElement("site",newID,sitesLong[ptDem["loc"]])										; need to transform site abbrevs
	wq.addElement("acct",newID,ptDem["loc"] ptDem["Account"])
	wq.addElement("ind",newID,ptDem["Indication"])
	wq.addElement("register",newID,{user:A_UserName},A_now)
	
	writeOut("/root/pending","enroll[@id='" id "']")
	
	return
}

moveHL7dem() {
	global fldVal, obxVal
	if (fldVal.acct) {	
		name := parseName(fldval.name)
		fldVal["dem-Name"] := fldval.Name
		fldVal["dem-Name_L"] := name.last
		fldVal["dem-Name_F"] := name.first
		fldVal["dem-MRN"] := fldval.MRN
		fldVal["dem-DOB"] := fldval.DOB
		fldVal["dem-Sex"] := fldval.Sex
		fldVal["dem-Indication"] := fldVal.ind
		fldVal["dem-Site"] := fldVal.site
		fldVal["dem-Billing"] := RegExReplace(fldVal.acct,"[[:alpha:]]")
		fldVal["dem-Ordering"] := fldval.prov
		fldval["dem-Device_SN"] := strX(fldval.dev," ",0,1,"",0,0)
	} else {
		fldVal["dem-Name"] := obxVal["PID_NameL"] ", " obxVal["PID_NameF"]
		fldVal["dem-Name_L"] := obxVal["PID_NameL"]
		fldVal["dem-Name_F"] := obxVal["PID_NameF"]
		fldVal["dem-MRN"] := obxVal["PID_PatMRN"]
		fldVal["dem-DOB"] := niceDate(obxVal["PID_DOB"])
		fldVal["dem-Sex"] := (obxVal["PID_Sex"]~="F") ? "Female" : "Male"
		fldVal["dem-Indication"] := obxVal.Diagnosis
		;~ fldVal["dem-Site"] := fldVal.site
		;~ fldVal["dem-Billing"] := RegExReplace(fldVal.acct,"[[:alpha:]]")
		fldVal["dem-Ordering"] := filterProv(obxVal["PV1_AttgNameF"] " " obxVal["PV1_AttgNameL"]).name
		;~ fldval["dem-Device_SN"] := strX(fldval.dev," ",0,1,"",0,0)
	}
	return
}

ProcessHl7PDF:
{
/*	Associate fldVal data with extra metadata from extracted PDF, complete final CSV report, handle files
*/
	fileNam := RegExReplace(fldVal.Filename,"i)\.pdf")									; fileNam is name only without extension, no path
	fileIn := hl7Dir fldVal.Filename													; fileIn has complete path \\childrens\files\HCCardiologyFiles\EP\HoltER Database\Holter PDFs\steve.pdf
	
	if (fileNam="") {																	; No PDF extracted
		eventlog("No PDF extracted.")
		MsgBox No PDF extracted!
		return
	}
	
	runwait, pdftotext.exe -l 2 "%fileIn%" "%filenam%.txt",,min							; convert PDF pages 1-2 with no tabular structure
	FileRead, newtxt, %filenam%.txt														; load into newtxt
	FileDelete, %filenam%.txt
	StringReplace, newtxt, newtxt, `r`n`r`n, `r`n, All									; remove double CRLF
	FileAppend % newtxt, %filenam%.txt													; create new tempfile with result, minus PDF
	FileMove %filenam%.txt, .\tempfiles\*, 1											; move a copy into tempfiles for troubleshooting
	FileAppend % fldval.hl7, %filenam%_hl7.txt											; create new tempfile with result, minus PDF
	FileMove %filenam%_hl7.txt, .\tempfiles\*, 1										; move a copy into tempfiles for troubleshooting
	
	type := fldval["OBR_TestCode"]														; study report type in OBR_testcode field
	if (type="Holter") {
		gosub Holter_Pr_Hl7
	} else if (type~="CEM|EOS") {
		gosub Event_BGH_Hl7
	} else {
		MsgBox % "No match!`n" type
		return
	}
	
	return
}

ProcessPDF:
{
/*	This main loop accepts a %fileIn% filename,
 *	determines the filetype based on header contents,
 *	concatenates the CSV strings of header (fileOut1) and values (fileOut2)
 *	into a single file (fileOut),
 *	move around the temp, CSV, and PDF files.
 */
	RunWait, pdftotext.exe -l 2 -table -fixed 3 "%fileIn%" "%filenam%.txt"				; convert PDF pages 1-2 to txt file
	newTxt:=""																			; clear the full txt variable
	FileRead, maintxt, %filenam%.txt													; load into maintxt
	FileDelete, %filenam%.txt
	StringReplace, newtxt, maintxt, `r`n`r`n, `r`n, All
	FileAppend %newtxt%, %filenam%.txt													; create new tempfile with newtxt result
	FileMove %filenam%.txt, .\tempfiles\%fileNam%.txt, 1								; move a copy into tempfiles for troubleshooting
		
	if (instr(newtxt,"zio xt")) {														; Processing loop based on identifying string in newtxt
		gosub Zio
	} else if (instr(newtxt,"Preventice") && instr(newtxt,"HScribe")) 	{				; New Preventice Holter 2017
		gosub Holter_Pr2
	} else if (instr(newtxt,"Preventice") && instr(newtxt,"End of Service Report")) {	; Body Guardian Heart CEM
		gosub Event_BGH
	;~ } else if (instr(newtxt,"Preventice") && instr(newtxt,"H3Plus")) {				; Original Preventice Holter
		;~ gosub Holter_Pr
	} else {
		eventlog(fileNam " bad file.")
		MsgBox No match!
		return
	}
	if (fetchQuit=true) {																; exited demographics fetchGUI
		return																			; so skip processing this file
	}
return
}

outputfiles:
{
	/*	Output the results and move files around
	*/
	fileOut1 := trim(fileOut1,",`t`r`n") "`n"												; make sure that there is only one `n 
	fileOut2 := trim(fileOut2,",`t`r`n") "`n"												; on the header and data lines
	fileout := fileOut1 . fileout2															; concatenate the header and data lines
	tmpDate := parseDate(fldval["dem-Test_Date"])											; get the study date from PDF result
	filenameOut := fldval["dem-MRN"] " " fldval["dem-Name_L"] " " tmpDate.MM "-" tmpDate.DD "-" tmpDate.YYYY
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
	FileDelete, % hl7dir fileNam ".hl7"														; We can delete the original HL7, if exists
	eventlog("Move files '" fileIn "' -> '" filenameOut)
	
	fileWQ := ma_date "," user "," 															; date processed and MA user
			. """" fldval["dem-Ordering"] """" ","														; extracted provider
			. """" fldval["dem-Name_L"] ", " fldval["dem-Name_F"] """" ","							; CIS name
			. """" fldval["dem-MRN"] """" ","													; CIS MRN
			. """" fldval["dem-Test_date"] """" ","											; extracted Test date (or CIS encounter date if none)
			. """" fldval["dem-Test_end"] """" ","											; extracted Test end
			. """" fldval["dem-Site"] """" ","												; CIS location
			. """" fldval["dem-Indication"] """" ","										; Indication
			. """" monType """" ; ","														; Monitor type
			. "`n"
	FileAppend, %fileWQ%, .\logs\fileWQ.csv													; Add to logs\fileWQ list
	FileCopy, .\logs\fileWQ.csv, %chipDir%fileWQ-copy.csv, 1
	
	wq := new XML("worklist.xml")
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

wqSetVal(id,node,val) {
	global wq
	
	newID := "/root/pending/enroll[@id='" id "']"
	k := wq.selectSingleNode(newID "/" node)
	
	if IsObject(k) {
		wq.setText(newID "/" node,val)
	} else {
		wq.addElement(node,newID,val)
	}
	
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
	if !(ptDem.EncDate) {																; must have a date to figure it out
		return
	}
	encDT := parseDate(ptDem.EncDate).YYYY . parseDate(ptDem.EncDate).MM . parseDate(ptDem.EncDate).DD
	inptMD := checkCrd(ptDem.Provider) 
	if (inptMD.fuzz<0.15) {																; attg is Crd
		ptDem.Loc := "Inpatient"														; set Loc so it won't fail
		return
	} 
	if (ymatch := y.selectSingleNode("//call[@date='" encDT "']/Ward_A").text) {		; found the Ward attg
		inptMD := checkCrd(ymatch)														; put in LAST, FIRST format
		if (inptMD.fuzz<0.15) {															; close enough match
			ptDem.Loc := "Inpatient"
			ptDem.Provider := inptMD.best
			eventlog("Cardiologist autoselected " ptDem.Provider )
			return
		}
	} else {
		gosub getMD																		; when all else fails, ask
		ptDem.Loc := "Inpatient"
	}
return
}

epRead() {
	global y, user, ma_date, fldval
	dlDate := A_Now
	FormatTime, dlDay, %dlDate%, dddd
	if (dlDay="Friday") {
		dlDate += 3, Days
	}
	FormatTime, dlDate, %dlDate%, yyyyMMdd
	
	RegExMatch(y.selectSingleNode("//call[@date='" dlDate "']/EP").text, "Oi)(Chun|Salerno|Seslar)", ymatch)
	if !(ep := ymatch.value()) {
		ep := cmsgbox("Electronic Forecast not complete","Which EP on Monday?","Chun|Salerno|Seslar","Q")
		eventlog("Reading EP assigned to " ep ".")
	}
	
	if (RegExMatch(fldval["dem-Ordering"], "Oi)(Chun|Salerno|Seslar)", epOrder))  {
		ep := epOrder.value()
	}
	
	FormatTime, ma_date, A_Now, MM/dd/yyyy
	fieldcoladd("","EP_read",ep)
	fieldcoladd("","EP_date",niceDate(dlDate))
	fieldcoladd("","MA",user)
	fieldcoladd("","MA_date",ma_date)
return
}

Holter_Pr_Hl7:
{
/*	Process newtxt from pdftotxt from HL7 extract
*/
	eventlog("Holter_Pr_HL7")
	monType := "PR"
	fullDisc := "i)60\s+s(ec)?/line"
	
	demog := stregX(newtxt,"Name:",1,0,"Medications:",1)
	fields[1] := ["Recording Start Date/Time","\R"
		, "Date Processed","(Technician|Hookup Tech)","Analyst","\R"
		, "Recording Duration","Recorder (No|Number)","\R"
		, "Indications","\R"]
	labels[1] := ["Test_date","null"
		, "Scan_date","Hookup_tech","Scanned_by","null"
		, "Recording_time","Device_SN","null"
		, "Indication","null"]
	fieldvals(demog,1,"dem")
	
	duration := stregx(newtxt "<<<","(\R)ALL BEATS",1,0,"(\R)HEART RATE EPISODES",0)
	fields[1] := ["Original Duration","Recording Duration","Analyzed Duration","Artifact Duration","\R"]
	labels[1] := ["null","Recording_time","Analysis_time","null","null"]
	fieldvals(duration,1,"dem")
	formatfield("dem","Test_end",fldval["dem-Recording_time"])
	
	progress, off
	
	if !(fldval.acct) {																	; fldval.acct exists if Holter has been processed
		gosub checkProc																	; get valid demographics
		if (fetchQuit=true) {
			return
		}
	}
	
	fieldsToCSV()
	
	FileGetSize, fileInSize, % hl7dir fldval.Filename
	if (fileInSize > 2000000) {															; probably a full disclosure PDF
		shortenPDF(fullDisc)															; generates .pdf and sh.pdf versions
	} 
	else loop {																			; just a short PDF
		if (findFullPDF(fldval.wqid)) {
			eventlog("Full disclosure PDF found.")
			break																		; found matching full disclosure, exit loop
		} else {
			eventlog("Full disclosure PDF not found.")
			
			MsgBox, 262164
				, Missing PDF
				, Full disclosure not found.`n`nSend request to Preventice?
			IfMsgBox, Yes
			{
				; send email
				return
			}
			
			MsgBox, 4149
				, Missing file
				, % "No full disclosure PDF found for:`n"
				. fldval["dem-Name_L"] ", " fldval["dem-Name_F"] "`n`n"
				. "Download PDF from ftp.eCardio.com website`n"
				. "then click [Retry] to proceed."
			IfMsgBox, Retry 
			{
				findFullPDF()															; rescan holterDir
				continue																; redo the loop
			} else {
				FileDelete, % fileIn
				eventlog("Refused to get full disclosure. Extracted PDF deleted.")
				Exit																	; either Cancel or X, go back to main GUI
			}
		}
	}
	
	fieldcoladd("","INTERP","")
	fieldcoladd("","Mon_type","Holter")
	
	fldval.done := true
	
return
}

fieldsToCSV() {
/*	tabs = tab-delim string
	"hrd-Total_beats(0)" -> fldval["hrd-Total_beats"] (default 0 if null)
	Regenerates new fileOut
*/
	global fldval, fileOut1, fileOut2, monType
	
	if (monType~="PR|ZIO") {
		tabs := "dem-Name_L	dem-Name_F	dem-Name_M	dem-MRN	dem-DOB	dem-Sex(NA)	dem-Site	dem-Billing	dem-Device_SN	dem-VOID1	"
			. "dem-Hookup_tech	dem-VOID2	dem-Meds	dem-Ordering	dem-Scanned_by	dem-Reading	"
			. "dem-Test_date	dem-Scan_date	dem-Hookup_time	dem-Recording_time	dem-Analysis_time	dem-Indication	dem-VOID3	"
			. "hrd-Total_beats(0)	hrd-Min(0)	hrd-Min_time	hrd-Avg(0)	hrd-Max(0)	hrd-Max_time	hrd-HRV	"
			. "ve-Total(0)	ve-Total_per(0)	ve-Runs(0)	ve-Beats(0)	ve-Longest(0)	ve-Longest_time	ve-Fastest(0)	ve-Fastest_time	"
			. "ve-Triplets(0)	ve-Couplets(0)	ve-SinglePVC(0)	ve-InterpPVC(0)	ve-R_on_T(0)	ve-SingleVE(0)	ve-LateVE(0)	"
			. "ve-Bigem(0)	ve-Trigem(0)	ve-SVE(0)	sve-Total(0)	sve-Total_per(0)	sve-Runs(0)	sve-Beats(0)	"
			. "sve-Longest(0)	sve-Longest_time	sve-Fastest(0)	sve-Fastest_time	sve-Pairs(0)	sve-Drop(0)	sve-Late(0)	"
			. "sve-LongRR(0)	sve-LongRR_time	sve-Single(0)	sve-Bigem(0)	sve-Trigem(0)	sve-AF(0)"
		;~ tabs := "dem-Name_L	dem-Name_F	dem-Name_M	dem-MRN	dem-DOB	dem-Sex(NA)	dem-Site	dem-Billing	dem-Device_SN	"			; for when we are full TRRIQ
			;~ . "dem-Hookup_tech	dem-Ordering	dem-Scanned_by	"
			;~ . "dem-Test_date	dem-Scan_date	dem-Hookup_time	dem-Recording_time	dem-Analysis_time	dem-Indication	"
			;~ . "hrd-Total_beats(0)	hrd-Min(0)	hrd-Min_time	hrd-Avg(0)	hrd-Max(0)	hrd-Max_time	hrd-HRV	"
			;~ . "ve-Total(0)	ve-Total_per(0)	ve-Runs(0)	ve-Beats(0)	ve-Longest(0)	ve-Longest_time	ve-Fastest(0)	ve-Fastest_time	"
			;~ . "ve-Triplets(0)	ve-Couplets(0)	ve-SinglePVC(0)	ve-InterpPVC(0)	ve-R_on_T(0)	ve-SingleVE(0)	ve-LateVE(0)	"
			;~ . "ve-Bigem(0)	ve-Trigem(0)	ve-SVE(0)	sve-Total(0)	sve-Total_per(0)	sve-Runs(0)	sve-Beats(0)	"
			;~ . "sve-Longest(0)	sve-Longest_time	sve-Fastest(0)	sve-Fastest_time	sve-Pairs(0)	sve-Drop(0)	sve-Late(0)	"
			;~ . "sve-LongRR(0)	sve-LongRR_time	sve-Single(0)	sve-Bigem(0)	sve-Trigem(0)	sve-AF(0)"
	} else if (monType="BGH") {
		tabs := "dem-Name_L	dem-Name_F	dem-MRN	dem-Ordering	dem-Sex(NA)	dem-DOB	dem-VOID_Practice dem-Indication	"
			. "dem-Test_date	dem-Test_end	dem-VOID	dem-Billing	"
			. "counts-Critical(0)	counts-Total(0)	counts-Serious(0)	counts-Manual(0)	counts-Stable(0)	counts-Auto(0)"
		;~ tabs := "dem-Name_L	dem-Name_F	dem-MRN	dem-DOB	dem-Sex(NA)	dem-Ordering	dem-Site	dem-Billing	dem-Device_SN	"		; for when we are full TRRIQ
			;~ . "dem-Test_date	dem-Test_end	dem-Indication	"
			;~ . "counts-Critical	counts-Total	counts-Serious	counts-Manual	counts-Stable	counts-Auto"
	}
	fileOut1 := ""
	fileOut2 := ""
	loop, parse, tabs, `t
	{
		x := A_LoopField																; PRE-LAB(default val)
		fld := strX(x,"",1,0,"(",1,1)													; field
			pre := strX(fld,"",1,0,"-",1,1)												; prefix
			lab := strX(fld,"-",1,1,"",0)												; label
		def := strX(x,"(",1,1,")",1,1)													; default value
		val := fldval[fld]																; value in fldval[pre-lab]
		res := (val = "") ? def : val													; result is value if exists, else default
		formatfield(pre,lab,res)														; sends formatted results, i.e. recreates fresh fileOut
	}
	eventlog("Fields mapping complete.")
	
return	
}

shortenPDF(find) {
	eventlog("ShortenPDF")
	global fileIn, fileNam, wincons
	sleep 500
	fullNam := filenam "full.txt"

	Progress,,,Scanning full size PDF...
	Runwait, pdftotext.exe "%fileIn%" "%fullnam%",,min,wincons								; convert PDF all pages to txt file
	eventlog("Extracting full text.")
	progress,100,, Shrinking PDF...
	FileRead, fulltxt, %fullnam%
	findpos := RegExMatch(fulltxt,find)
	pgpos := instr(fulltxt,"Page ",,findpos-strlen(fulltxt))
	RegExMatch(fulltxt,"Oi)Page\s+(\d+)\s",pgs,pgpos)
	pgpos := pgs.value(1)
	RunWait, pdftk.exe "%fileIn%" cat 1-%pgpos% output "%fileIn%sh.pdf",,min
	if !FileExist(fileIn "sh.pdf") {
		FileCopy, %fileIn%, %fileIn%sh.pdf
	}
	filedelete, %fullnam%
	FileGetSize, sizeIn, %fileIn%
	FileGetSize, sizeOut, %fileIn%sh.pdf
	eventlog("IN: " thousandsSep(sizeIn) ", OUT: " thousandsSep(sizeOut))
	progress, off
return	
}

findFullPdf(wqid:="") {
/*	Scans HolterDir for potential full disclosure PDFs
	maybe rename if appropriate
*/
	global holterDir, hl7dir, fldval, pdfList
	
	pdfList := Object()																	; clear list to add to WQlist
	
	fileCount := ComObjCreate("Scripting.FileSystemObject").GetFolder(holterDir).Files.Count
	
	Loop, files, %holterDir%*.pdf
	{
		fname := A_LoopFileName, fileIn := A_LoopFileFullPath, fnam := RegExReplace(fname,"i)\.pdf")
		progress, % 100*A_index/fileCount, % fname, Scanning PDFs folder
		
		;---Skip any PDFs that have already been processed or are in the middle of being processed
		if (fname~="i)(sh|-short)\.pdf") 
			continue
		if FileExist(fname "sh.pdf") 
			continue
		if FileExist(fnam "-short.pdf") 
			continue
		
		RegExMatch(fname,"O)_WQ(\d+)(\w)?\.pdf",fnID)									; get filename WQID if PDF has already been renamed
		
		if (readWQ(fnID.1).node = "done") {
			eventlog("Leftover PDF: " fnam ", moved to archive.")
			FileMove, % fileIn, % holterDir "archive\" fname
			continue
		}
		
		if (fnID.0 = "") {																; Unprocessed full disclosure PDF
			runwait, pdftotext.exe -l 1 "%fileIn%" "%fnam%.txt",,min					; convert PDF pages 1-2 with no tabular structure
			FileRead, newtxt, %fnam%.txt												; load into newtxt
			FileDelete, %fnam%.txt
			StringReplace, newtxt, newtxt, `r`n`r`n, `r`n, All							; remove double CRLF
			
			flds := getPdfID(newtxt)
			
			newFnam := strQ(flds.nameL,"###_" flds.mrn,fnam) strQ(flds.wqid,"_WQ###")
			FileMove, %fileIn%, % holterDir newFnam ".pdf", 1							; rename the unprocessed PDF
			eventlog("Holter PDF: " fName " renamed to " newFnam)
			fName := newFnam ".pdf"
		} 
		pdfList.push(fName)
		
		if (wqid = "") {																; this is just a refresh loop
			continue																	; just build the list
		}
		
		if (fnID.1 == wqid) {															; filename WQID matches wqid arg
			FileMove, % hl7dir fldval.Filename, % hl7dir fldval.Filename "sh.pdf"		; rename the pdf in hl7dir to -short.pdf
			FileMove, % holterDir fName , % hl7dir fldval.filename 						; move this full disclosure PDF into hl7dir
			progress, off
			eventlog(fName " moved to hl7dir.")
			return true																	; stop search and return
		} else {
			continue
		}
	}
	progress, off
	return false																		; fell through without a match
}

getPdfID(txt) {
/*	Parses txt for demographics
	returns type=H,E,Z,M and demographics in an array, and wqid if found
	or error if no match
*/
	global fldval
	res := Object()
	
	if instr(txt,"MORTARA") {															; Mortara Holter
		res.type := "H"
		name := parseName(res.name := trim(stregX(txt,"Name:",1,1,"Recording Start",1)))
			res.nameL := name.last
			res.nameF := name.first
		dt := parseDate(trim(stregX(txt,"Start Date/Time:?",1,1,"\R",1)))
			res.date := dt.yyyy dt.mm dt.dd
			res.time := dt.hr dt.min
		dobDt := parseDate(trim(stregX(txt,"(Date of Birth|DOB):?",1,1,"\R",1)))
			res.dob := dobDt.yyyy dobDt.mm dobDt.dd
		res.mrn := trim(stregX(txt,"Secondary ID:?",1,1,"Age:?",1))
		res.ser := trim(stregX(txt,"Recorder (No|Number):?",1,1,"\R",1))
		res.wqid := strQ(findWQid(res.date,res.mrn).id,"###","00000") "H"
	} else if instr(txt,"BodyGuardian Heart") {
		res.type := "E"
		name := parseName(res.name := trim(stregX(txt,"Patient:",1,1,"Patient ID",1)," `t`r`n"))
			res.nameL := name.last
			res.nameF := name.first
		dt := parseDate(trim(stregX(txt,"Period \(.*?\R",1,1," - ",1)," `t`r`n"))
			res.date := dt.yyyy dt.mm dt.dd
		res.mrn := trim(stregX(txt,"Patient ID",1,1,"Gender",1)," `t`r`n")
		res.wqid := strQ(findWQid(res.date,res.mrn).id,"###E","00000E")
	} else if instr(txt,"Zio XT") {
		res.type := "Z"
		name := parseName(res.name := trim(stregX(txt,"Final Report for",1,1,"Date of Birth",1)," `t`r`n"))
			res.nameL := name.last
			res.nameF := name.first
		enroll := stregX(txt,"Enrollment Period",1,0,"Analysis Time",1)
		dt := parseDate(stregX(enroll,"i)\R+.*?(hours|days).*?\R+",1,1,",",1))
			res.date := dt.yyyy dt.mm dt.dd
		res.mrn := strQ(trim(stregX(txt,"Patient ID",1,1,"Managing Location",1)," `t`r`n"),"###","Zio")
		res.wqid := "00000Z"
	}
	return res
}

Holter_Pr2:
{
	eventlog("Holter_Pr2")
	monType := "PR"
	fullDisc := "i)60\s+s(ec)?/line"
	
	if (fileinsize < 2000000) {															; Shortened files are usually < 1-2 Meg
		eventlog("Filesize predicts non-full disclosure PDF.")							; Full disclosure are usually ~ 9-19 Meg
		MsgBox, 4112, Filesize error!, This file does not appear to be a full-disclosure PDF. Please download the proper file and try again.
		fetchQuit := true
		return
	}
	
	/* Pulls text between field[n] and field[n+1], place in labels[n] name, with prefix "dem-" etc.
	 */
	demog := stregX(newtxt,"Name:",1,0,"Conclusions",1)
	fields[1] := ["Name","\R","Recording Start Date/Time","\R"
		, "ID","Secondary ID","Admission ID","\R"
		, "Date Of Birth","Age","Gender","\R"
		, "Date Processed","(Referring|Ordering) Phys(ician)?","\R"
		, "Technician|Hookup Tech","Recording Duration","\R"
		, "Analyst","Recorder (No|Number)","\R"
		, "Indications","Medications","\R"]
	labels[1] := ["Name","null","Test_date","null"
		, "null","MRN","null","null"
		, "DOB","VOID_Age","Sex","null"
		, "Scan_date","Ordering","null"
		, "Hookup_tech","VOID_Duration","null"
		, "Scanned_by","Device_SN","null"
		, "Indication","VOID_meds","null"]
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
	fldVal["dem-Test_end"] := RegExReplace(fldVal["dem-Recording_time"],"(\d{1,2}) hr (\d{1,2}) min","$1:$2")	; Places value for fileWQ
	
	rateStat := stregX(sumStat,"VENTRICULAR ECTOPY",1,0,"PACED|SUPRAVENTRICULAR ECTOPY",1)
	fields[2] := ["Ventricular Beats","Singlets","Couplets","Runs","Fastest Run","Longest Run","R on T Beats"]
	labels[2] := ["Total","SingleVE","Couplets","Runs","Fastest","Longest","R on T"]
	scanParams(rateStat,2,"ve",1)
	
	rateStat := stregX(sumStat "<<<","SUPRAVENTRICULAR ECTOPY",1,0,"<<<|OTHER RHYTHM EPISODES",1)
	fields[3] := ["Supraventricular Beats","Singlets","Pairs","Runs","Fastest Run","Longest Run"]
	labels[3] := ["Total","Single","Pairs","Runs","Fastest","Longest"]
	scanParams(rateStat,3,"sve",1)
	
	gosub checkProc												; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	fieldsToCSV()
	tmpstr := stregx(newtxt,"Conclusions",1,1,"Reviewing Physician",1)
	StringReplace, tmpstr, tmpstr, `r, `n, ALL
	fieldcoladd("","INTERP",trim(cleanspace(tempstr)," `n"))
	fieldcoladd("","Mon_type","Holter")
	
	ShortenPDF(fullDisc)
	
	fldval.done := true

return
}

CheckProc:
{
	eventlog("CheckProcPr")
	fetchQuit := false
	
	if (fldval.node = "done") {
	;~ if (zzzfldval.node = "done") {
		MsgBox % fileIn " has been scanned already.`n`nDeleting file."
		eventlog(fileIn " already scanned. PDF deleted.")
		FileDelete, % fileIn
		fetchQuit := true
		return
	}
	
	ptDem := Object()																	; Populate temp object ptDem with parsed data from PDF fldVal
	ptDem["nameL"] := fldVal["dem-Name_L"]
	ptDem["nameF"] := fldVal["dem-Name_F"] 
	ptDem["Name"] := fldval["dem-Name"]
	ptDem["mrn"] := fldVal["dem-MRN"] 
	ptDem["DOB"] := fldVal["dem-DOB"] 
	ptDem["Sex"] := fldVal["dem-Sex"]
	ptDem["Loc"] := fldVal["dem-Site"]
	ptDem["Account"] := fldVal["dem-Billing"]											; If want to force click, don't include Acct Num
	ptDem["Provider"] := filterProv(fldVal["dem-Ordering"]).name
	ptDem["EncDate"] := fldVal["dem-Test_date"]
	ptDem["Indication"] := fldVal["dem-Indication"]
	eventlog("PDF demog: " ptDem.nameL ", " ptDem.nameF " " ptDem.mrn " " ptDem.EncDate)
	
	if (fldval.acct) {																	; <acct> exists, has been registered or uploaded through TRRIQ
		eventlog("Pulled valid data for " fldval.name " " fldval.mrn " " fldval.date)
		MsgBox, 4160, Found valid registration, % "" 
		  . fldval.name "`n" 
		  . "MRN " fldval.mrn "`n" 
		  . "Acct " fldval.acct "`n" 
		  . "Ordering: " fldval.prov "`n" 
		  . "Study date: " fldval.date "`n`n" 
	} 
	else {
	;~ else if false {
		/*	Did not return based on done or valid status, 
		 *	and has not been validated yet so no prior TRRIQ data
		 */
		Clipboard := ptDem.nameL ", " ptDem.nameF											; fill clipboard with name, so can just paste into CIS search bar
		MsgBox, 4096,, % "Extracted data for:`n"
			. "   " ptDem.nameL ", " ptDem.nameF "`n"
			. "   " ptDem.mrn "`n"
			. "   " ptDem.EncDate "`n`n"
			. "Paste clipboard into CIS search to select patient and encounter"
		
		gosub fetchGUI
		gosub fetchDem
		checkFetchDem(fldVal["dem-Name_L"],fldVal["dem-Name_F"],fldVal["dem-MRN"])			; make sure grabbed name (ptDem) matches PDF (fldVal)
		if (fetchQuit=true) {
			return
		}
		/*	When fetchDem successfully completes,
		 *	replace fldVal with newly acquired values
		 */
		fldVal.Name := ptDem["nameL"] ", " ptDem["nameF"]
		fldVal["dem-Name_L"] := fldval["Name_L"] := ptDem["nameL"]
		fldVal["dem-Name_F"] := fldval["Name_F"] := ptDem["nameF"] 
		fldVal["dem-MRN"] := ptDem["mrn"] 
		fldVal["dem-DOB"] := ptDem["DOB"] 
		fldVal["dem-Sex"] := ptDem["Sex"]
		fldVal["dem-Site"] := ptDem["Loc"]
		fldVal["dem-Billing"] := ptDem["Account"]
		fldVal["dem-Ordering"] := ptDem["Provider"]
		fldVal["dem-Test_date"] := ptDem["EncDate"]
		fldVal["dem-Indication"] := ptDem["Indication"]
		
		filecheck()
		FileOpen(".lock", "W")																; Create lock file.
			if (fldval.wqid) {
				id := fldval.wqid
			} else {
				id := A_TickCount 															; create wqid record if it doesn't exist somehow
				wq.addElement("enroll","/root/pending",{id:id})
			}
			newID := "/root/pending/enroll[@id='" id "']"
			wqSetVal(id,"date",(ptDem["date"]) ? ptDem["date"] : substr(A_now,1,8))
			wqSetVal(id,"name",ptDem["nameL"] ", " ptDem["nameF"])
			wqSetVal(id,"mrn",ptDem["mrn"])
			wqSetVal(id,"sex",ptDem["Sex"])
			wqSetVal(id,"dob",ptDem["dob"])
			wqSetVal(id,"dev"
				, (montype="PR" ? "Mortara H3+ - " 
				: montype="BGH" ? "BodyGuardian Heart - BG"
				: montype="ZIO" ? "Zio" : "")
				. fldVal["dem-Device_SN"])
			wqSetVal(id,"prov",ptDem["Provider"])
			wqSetVal(id,"site",sitesLong[ptDem["loc"]])										; need to transform site abbrevs
			wqSetVal(id,"acct",ptDem["loc"] ptDem["Account"])
			wqSetVal(id,"ind",ptDem["Indication"])
		filedelete, .lock
		writeOut("/root/pending","enroll[@id='" id "']")
		
		eventlog("Demographics updated for WQID " fldval.wqid ".") 
	}
	
	;---Copy ptDem back to fldVal, whether fetched or not
	fldVal.Name := ptDem["nameL"] ", " ptDem["nameF"]
	fldVal["dem-Name_L"] := fldval["Name_L"] := ptDem["nameL"]
	fldVal["dem-Name_F"] := fldval["Name_F"] := ptDem["nameF"] 
	fldVal["dem-MRN"] := fldval["MRN"] := ptDem["mrn"] 
	fldVal["dem-DOB"] := ptDem["DOB"] 
	fldVal["dem-Sex"] := ptDem["Sex"]
	fldVal["dem-Site"] := ptDem["Loc"]
	fldVal["dem-Billing"] ptDem["Account"]
	fldVal["dem-Ordering"] := ptDem["Provider"]
	fldVal["dem-Test_date"] := ptDem["EncDate"]
	fldVal["dem-Indication"] := ptDem["Indication"]
	
return
}

Zio:
{
	eventlog("Holter_Zio")
	monType := "Zio"
	
	RunWait, pdftotext.exe -table -fixed 3 "%fileIn%" "%filenam%.txt", , hide			; reconvert entire Zio PDF 
	newTxt:=""																		; clear the full txt variable
	FileRead, maintxt, %filenam%.txt												; load into maintxt
	StringReplace, newtxt, maintxt, `r`n`r`n, `r`n, All
	StringReplace, newtxt, newtxt, % chr(12), >>>page>>>`r`n, All
	FileDelete %filenam%.txt														; remove any leftover tempfile
	FileAppend %newtxt%, %filenam%.txt												; create new tempfile with newtxt result
	FileMove %filenam%.txt, .\tempfiles\%fileNam%.txt, 1							; overwrite copy in tempfiles
	eventlog("Zio PDF rescanned -> " fileNam ".txt")
	
	zcol := columns(newtxt,"","SIGNATURE",0,"Enrollment Period") ">>>end"
	demo1 := onecol(cleanblank(stregX(zcol,"\s+Date of Birth",1,0,"Prescribing Clinician",1)))
	demo2 := onecol(cleanblank(stregX(zcol,"\s+Prescribing Clinician",1,0,"\s+(Supraventricular Tachycardia \(|Ventricular tachycardia \(|AV Block \(|Pauses \(|Atrial Fibrillation)",1)))
	demog := RegExReplace(demo1 "`n" demo2,">>>end") ">>>end"
	
	znam := strVal(zcol,"Report for","Date of Birth")
	formatfield("dem","Name",znam)
	
	fields[1] := ["Date of Birth","Patient ID","Gender","Primary Indication","Prescribing Clinician","(Referring Clinician|Managing Location)",">>>end"]
	labels[1] := ["DOB","MRN","Sex","Indication","Ordering","Site","end"]
	fieldvals(demog,1,"dem")
	
	tmp := oneCol(stregX(zcol,"Enrollment Period",1,0,"Heart\s+Rate",1))
	enroll := strVal(tmp,"Enrollment Period","Analysis Time")
	fieldcoladd("dem","Test_date",strVal(enroll,"hours?",","))
	fieldcoladd("dem","Test_end",strVal(enroll,"to\s",","))
	fieldcoladd("dem","Analysis_time",strVal(tmp,"Analysis Time","\(after"))
	fieldcoladd("dem","Recording_time",strVal(tmp,"Enrollment Period","\R\d{1,2}/\d{1,2}"))

	znums := columns(zcol ">>>end","Enrollment Period",">>>end",1)
	
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
	
	gosub checkProc												; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	fieldsToCSV()
	
	zinterp := cleanspace(columns(newtxt,"Preliminary Findings","SIGNATURE",,"Final Interpretation"))
	zinterp := trim(StrX(zinterp,"",1,0,"Final Interpretation",1,20))
	
	fieldcoladd("","INTERP",zinterp)
	fieldcoladd("","Mon_type","Holter")
	
	FileCopy, %fileIn%, %fileIn%sh.pdf
	
	fldval.done := true

return
}

ZioArrField(txt,fld) {
	str := stregX(txt,fld,1,0,"#####",1)
	if instr(str,"Episodes") {
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

hl7fld(bl,pre) {
	global fields, labels, fldval
	
	for k, i in fields[bl]																; Step through each val "i" from fields[bl,k]
	{
		lbl := labels[bl][k]
		res := fldval[i]
		formatField(pre,lbl,res)
	}
	return
}

Event_BGH_Hl7:
{
	eventlog("Event_BGH_HL7")
	monType := "BGH"
	
	fieldcoladd("dem","Test_date",niceDate(obxVal["Enroll_Start_Dt"]))
	fieldcoladd("dem","Test_end",niceDate(obxVal["Enroll_End_Dt"]))
	
	count:=[]																			; create object for counts
	count["Patient-Activated"]:=0														; zero the results instead of null
	count["Auto-Detected"]:=0
	count["Stable"]:=0
	count["Serious"]:=0
	count["Critical"]:=0
	for key,val in obxVal																; recurse through obxVal results
	{
		if (key~="Event_Acuity|Event_Type") {											; count Critical/Serious/Stable and Auto/Manual events
			count[val] ++																; more reliable than parsing PDF
		}
	}
	fieldcoladd("counts","Critical",count["Critical"])
	fieldcoladd("counts","Serious",count["Serious"])
	fieldcoladd("counts","Stable",count["Stable"])
	fieldcoladd("counts","Manual",count["Patient-Activated"])
	fieldcoladd("counts","Auto",count["Auto-Detected"])
	fieldcoladd("counts","Total",count["Auto-Detected"]+count["Patient-Activated"])
	
	gosub checkProc												; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	fieldstoCSV()
	
	fieldcoladd("","Mon_type","Event")
	
	FileCopy, %fileIn%, %fileIn%sh.pdf
	
	fldval.done := true
	
Return
}

Event_BGH:
{
	eventlog("Event_BGH")
	monType := "BGH"
	
	name := "Patient Name:   " trim(columns(newtxt,"Patient:","Enrollment Info",1,"")," `n")
	demog := columns(newtxt,"","(Summarized Findings|Event Summary|Rhythm Summary)",,"Enrollment Info")
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
	
	fields[1] := ["Patient Name", "Patient ID", "Physician", "Gender", "Date of Birth", "Practice", "Diagnosis"]
	labels[1] := ["Name", "MRN", "Ordering", "Sex", "DOB", "VOID_Practice", "Indication"]
	fieldvals(demog,1,"dem")
	fldval["name_L"] := ptDem["nameL"]
	
	tmpDT := strVal(enroll,"Period \(.*\)","Event Counts")									; Study date
	fieldcoladd("dem","Test_date",trim(strX(tmpDT,"",1,1," ",1,1)," `r`n"))
	fieldcoladd("dem","Test_end",trim(strX(tmpDT," - ",0,3,"",0)," `r`n"))
	
	fields[3] := ["Critical","Total","Serious","(Manual|Pt Trigger)","Stable","Auto Trigger"]
	labels[3] := ["Critical","Total","Serious","Manual","Stable","Auto"]
	fieldvals(enroll,3,"counts")
	
	gosub checkProc												; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	fieldstoCSV()
	
	fieldcoladd("","Mon_type","Event")
	
	FileCopy, %fileIn%, %fileIn%sh.pdf
	
	fldval.done := true
	
Return
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
		;~ fldval[lbl] := m
		
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

	return trim(str.value("res")," :`n`r`t")
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
	
	if (lab="Name") {
		txt := RegExReplace(txt,"i),?( JR| III| IV)$")						; Filter out name suffixes
		name := parseName(txt)
		fieldColAdd(pre,"Name",name.last ", " name.first)
		fieldColAdd(pre,"Name_L",name.last)
		fieldColAdd(pre,"Name_F",name.first)
		return
	}
	if (lab="DOB") {														; remove (age) from DOB
		txt := strX(txt,"",1,0," (",2)
	}

	if (lab~="^(Referring|Ordering)$") {
		tmpCrd := checkCrd(txt)												; Get Crd, Grp, and Eml via checkCrd()
		fieldColAdd(pre,lab,tmpCrd.best)
		fieldColAdd(pre,lab "_grp",tmpCrd.group)
		fieldColAdd(pre,lab "_eml",Docs[tmpCrd.Group ".eml",ObjHasValue(Docs[tmpCrd.Group],tmpCrd.best)])
		if (tmpCrd="") {
			eventlog("*** Blank Crd value ***")
		}
		return
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
			fieldColAdd(pre,lab,txt)
			return
		}
		if (lab ~= "i)_time") {															; Any other _Time field, remove the date
			txt := parseDate(txt).time
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
	global fileOut1, fileOut2, fldVal
	pre := (pre="") ? "" : pre "-"
	if instr(fileOut1,"""" pre lab """") {
		return
	}
	fileOut1 .= """" pre lab ""","
	fileOut2 .= """" txt ""","
	fldVal[pre lab] := txt
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
	x := filterprov(x).name
	for rowidx,row in Docs																; Groups
	{
		if (substr(rowIdx,-3)=".eml") {
			continue
		}
		for colidx,item in row															; Providers
		{
			if (item="") {                                ; empty field will break fuzzysearch 
				continue 
			} 
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

checkFetchDem(nameL,nameF,mrn) {
/*	Check if fetchDem NAME and MRN match that parsed from PDF
	nameL, nameF, mrn all required params
	If bad match, returns fetchQuit=true
	If acceptable, returns fetchQuit=false
*/
	global ptDem, fetchQuit
	fullName := nameL ", " nameF
	fullNameDem := ptDem["nameL"] ", " ptDem["nameF"]
	fuzz := fuzzysearch(format("{:U}",fullName), format("{:U}",fullNameDem))
	thresh := 0.15													
	
	if (fuzz > thresh) {
		eventlog("Name fuzz error. "
			. "Parsed """ mrn """, """ fullName """ "
			. "Grabbed """ ptDem["mrn"] """, """ fullNameDem """.")
			
		if (mrn=ptDem["mrn"]) {															; correct MRN but bad name match
			MsgBox, 262193, % "Name error (" round((1-tmp)*100,2) "%)"
				, % "Name does not match!`n`n"
				.	"	Parsed:	" fullName "`n"
				.	"	Grabbed:	" fullNameDem "`n`n"
				.	"OK = use " fullNameDem "`n`n"										; "OK" will accept this fetchDem data
				.	"Cancel = skip this file"
			IfMsgBox, Cancel
			{
				eventlog("Cancel this file.")
				fetchQuit:=true															; cancel out of processing file
				return
			}
		} else {																		; just plain doesn't match
			MsgBox, 262160, % "Name error (" round((1-tmp)*100,2) "%)"
				, % "Name does not match!`n`n"
				.	"	Parsed:	" fullName "`n"
				.	"	Grabbed:	" fullNameDem "`n`n"
				.	"Skipping this file."
				
			eventlog("Demographics mismatch.")
			fetchQuit:=true
			return
		}
	}
	
	return
}

filterProv(x) {
/*	Filters out all irregularities and common typos in Provider name from manual entry
	Returns as {name:"Albers, Erin", site:"CRB"}
	Provider-Site may be in error
*/
	global sites, sites0
	
	allsites := sites "|" sites0
	RegExMatch(x,"i)-(" allsites ")\s*,",site)
	x := trim(x)																		; trim leading and trailing spaces
	x := RegExReplace(x,"i)\s*-\s*(" allsites ")$")										; remove trailing "LOUAY TONI(-tri)"
	x := RegExReplace(x,"i)( [a-z](\.)? )"," ")											; remove middle initial "STEPHEN P SESLAR" to "Stephen Seslar"
	x := RegExReplace(x,"i)^Dr(\.)?(\s)?")												; remove preceding "(Dr. )Veronica..."
	x := RegExReplace(x,"i)^[a-z](\.)?\s")												; remove preceding "(P. )Ruggerie, Dennis"
	x := RegExReplace(x,"i)\s[a-z](\.)?$")												; remove trailing "Ruggerie, Dennis( P.)"
	x := RegExReplace(x,"i)\s*-\s*(" allsites ")\s*,",",")								; remove "SCHMER(-YAKIMA), VERONICA"
	x := RegExReplace(x,"i) (MD|DO)$")													; remove trailing "( MD)"
	x := RegExReplace(x,"i) (MD|DO),",",")												; replace "Ruggerie MD, Dennis" with "Ruggerie, Dennis"
	StringUpper,x,x,T																	; convert "RUGGERIE, DENNIS" to "Ruggerie, Dennis"
	if !instr(x,", ") {
		x := strX(x," ",1,1,"",1,0) ", " strX(x,"",1,1," ",1,1)							; convert "DENNIS RUGGERIE" to "RUGGERIE, DENNIS"
	}
	x := RegExReplace(x,"^, ")															; remove preceding "(, )Albers" in event this happens
	if (site1="TRI") {																	; sometimes site improperly registered as "tri"
		site1 := "TRI-CITIES"
	}
	return {name:x, site:site1}
}

IEGet(name="") {
/*	from the very helpful post by jethrow
	https://autohotkey.com/board/topic/47052-basic-webpage-controls-with-javascript-com-tutorial/
*/
	IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame     ;// Get active window if no parameter
	Name := (Name="New Tab - Windows Internet Explorer")? "about:Tabs":RegExReplace(Name, " - (Windows|Microsoft)? ?Internet Explorer$")
	for wb in ComObjCreate("Shell.Application").Windows()
		if wb.LocationName=Name and InStr(wb.FullName, "iexplore.exe")
			return wb
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

strQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	return (var1="") ? null : RegExReplace(txt,"###",var1)
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

ParseName(x) {
/*	Determine first and last name
*/
	x := trim(x)																		; trim edges
	x := RegExReplace(x," \w "," ")														; remove middle initial: Troy A Johnson => Troy Johnson
	x := RegExReplace(x,"i),?( JR| III| IV)$")											; Filter out name suffixes
	x := RegExReplace(x,"\s+"," ",ct)													; Count " "
	
	if instr(x,",") 																	; Last, First
	{
		last := trim(strX(x,"",1,0,",",1,1))
		first := trim(strX(x,",",1,1,"",0))
	}
	else if (ct=1)																		; First Last
	{
		first := strX(x,"",1,0," ",1)
		last := strX(x," ",1,1,"",0)
	}
	else																				; James Jacob Jingleheimer Schmidt
	{
		x0 := x																			; make a copy to disassemble
		n := 1
		Loop
		{
			x0 := strX(x0," ",n,1,"",0)													; cut from first " " to end
			if (x0="") {
				q := trim(q,"|")
				break
			}
			q .= x0 "|"																	; add to button q
		}
		last := cmsgbox("Name check",x "`n" RegExReplace(x,".","--") "`nWhat is the patient's`nLAST NAME?",q)
		if (last~="close|xClose") {
			return {first:"",last:x}
		}
		first := RegExReplace(x," " last)
	}
	
	return {first:first,last:last}
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
	else if RegExMatch(x,"\b(\d{4})-?(\d{2})-?(\d{2})\b",d) {								; 20150103 or 2015-01-03
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

readIni(section) {
/*	Reads a set of variables
	[section]					==	 		var1 := res1, var2 := res2
	var1=res1
	var2=res2
	
	[array]						==			array := ["ccc","bbb","aaa"]
	=ccc
	bbb
	=aaa
	
	[objet]						==	 		objet := {aaa:10,bbb:27,ccc:31}
	aaa:10
	bbb:27
	ccc:31
*/
	global
	local x, i, key, val
		, i_res := object()
		, i_type := []
		, i_lines := []
	i_type.var := i_type.obj := i_type.arr := false
	IniRead,x,trriq.ini,%section%
	Loop, parse, x, `n,`r																; analyze section struction
	{
		i := A_LoopField
		if (i~="(?<!"")[=]")															; find = not preceded by "
		{
			if (i ~= "^=") {															; starts with "=" is an array list
				i_type.arr := true
			} else {																	; "aaa=123" is a var declaration
				i_type.var := true
			}
		} else																			; does not contain a quoted =
		{
			if (i~="(?<!"")[:]") {														; find : not preceded by " is an object
				i_type.obj := true
			} else {																	; contains neither = nor : can be an array list
				i_type.arr := true
			}
		}
	}
	if ((i_type.obj) + (i_type.arr) + (i_type.var)) > 1 {								; too many types, return error
		return error
	}
	Loop, parse, x, `n,`r																; now loop through lines
	{
		i := A_LoopField
		if (i_type.var) {
			key := strX(i,"",1,0,"=",1,1)
			val := strX(i,"=",1,1,"",0)
			%key% := trim(val,"""")
		}
		if (i_type.obj) {
			key := strX(i,"",1,0,":",1,1)
			val := strX(i,":",1,1,"",0)
			i_res[key] := trim(val,"""")
		}
		if (i_type.arr) {
			i := RegExReplace(i,"^=")													; remove preceding =
			i_res.push(trim(i,""""))
		}
	}
	return i_res
}

#Include CMsgBox.ahk
#Include xml.ahk
#Include sift3.ahk
#Include hl7.ahk
