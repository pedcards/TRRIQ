/*	TRRIQ - The Rhythm Recording Interpretation Query
	Disassembles HL7 and PDF files into discrete data elements
	Outputs into a format readable by HolterDB (CSV, PDF, and short PDF)
	Sends report to HIM
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force  ; only allow one running instance per user
#MaxMem 128
#Include %A_ScriptDir%
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

progress,,% " ",TRRIQ intializing...
FileGetTime, wqfileDT, .\files\wqupdate
FileGetTime, runningVer, % A_ScriptName

SplitPath, A_ScriptDir,,fileDir
user := A_UserName
userinstance := substr(tobase(A_TickCount,36),-3)
IfInString, fileDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	isDevt := true
	path:=readIni("devtpaths")
	eventlog(">>>>> Started in DEVT mode.")
} else {
	isDevt := false
	path:=readIni("paths")
	eventlog(">>>>> Started in PROD mode. " A_ScriptName " ver " substr(runningVer,1,12) " " A_Args[1])
}
if (A_Args[1]~="launch") {
	eventlog("***** launched from legacy shortcut.")
}

readini("setup")

/*	Get location info
*/
#Include HostName.ahk
progress,,% " ",Identifying workstation...
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
wksVoid := StrSplit(wksVM, "|")

/*	Get valid WebUploadDir
*/
webUploadDir := check_h3(path.webupload,webUploadStr)									; Find the location of H3 data files
checkMachine()

/*	Read outdocs.csv for Cardiologist and Fellow names 
*/
progress,,% " ",Scanning providers...
Docs := Object()
tmpChk := false
if FileExist(path.chip "outdocs.csv") {													; if server access to chipotle outdocs, make a local copy
	FileCopy, % path.chip "outdocs.csv", .\files\outdocs.csv, 1
}
Loop, Read, .\files\outdocs.csv
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

/*	Generate worklist.xml if missing
*/
if fileexist("worklist.xml") {
	wq := new XML("worklist.xml")
} else {
	wq := new XML("<root/>")
	wq.addElement("pending","/root")
	wq.addElement("done","/root")
	wq.save("worklist.xml")
}

/*	Read call schedule (Electronic Forecast and Qgenda)
*/
fcVals := readIni("Forecast")
Forecast_svc := []
Forecast_val := []
for key,val in fcVals
{
	tmpVal := strX(val,"",1,0,":",1)
	tmpStr := strX(val,":",1,1,"",0)
	Forecast_svc.Insert(tmpVal)
	Forecast_val.Insert(tmpStr)
}
updateCall()

/*	Initialize rest of vars and strings
*/
Progress, , % " ", Initializing variables
demVals := readIni("demVals")																		; valid field names for parseClip()

indCodes := readIni("indCodes")
for key,val in indCodes
{
	tmpVal := strX(val,"",1,0,":",1)
	tmpStr := strX(val,":",1,1,"",0)
	indOpts .= tmpStr "|"
}

monCodes := readIni("EpicMonitorType")

initHL7()
hl7DirMap := {}

Progress, , % " ", Cleaning old .bak files
Loop, files, bak\*.bak
{
	dt := A_now
	dt -= RegExReplace(A_LoopFileName,"\.bak"), Days
	if (dt > 7) {
		FileDelete, bak\%A_LoopFileName%
	}
}
Progress, Off

MainLoop: ; ===================== This is the main part ====================================
{
	Loop
	{
		Gosub PhaseGUI
		WinWaitClose, TRRIQ Dashboard
		
		if (phase="Upload") {
			eventlog("Start Mortara upload.")
			mwuPhase := "Transfer"
			MortaraUpload(mwuPhase)
		}
	}
	
	checkPreventiceOrdersOut()
	cleanDone()
	
	ExitApp
}

PhaseGUI:
{
	phase :=
	Gui, phase:Destroy
	Gui, phase:Default
	Gui, +AlwaysOnTop

	Gui, Add, Text, x670 y15 w200 vPhaseNumbers, "`n`n"
	Gui, Add, GroupBox, x660 y0 w220 h65
	
	Gui, Font, Bold
	Gui, Add, Button
		, Y+10 wp h40 gPhaseRefresh
		, Refresh lists
	Gui, Add, Button
		, Y+10 wp h40 gPrevGrab Disabled
		, Grab Preventice enrollments
	Gui, Add, Text, wp h40																; space between top buttons and lower buttons
	Gui, Add, Text, Y+10 wp h24 Center, Register/Prepare a`nHOLTER or EVENT MONITOR
	Gui, Add, Button
		, Y+10 wp h40 vRegister gPhaseOrder DISABLED
		, No active orders
	Gui, Add, Text, wp h30
	Gui, Add, Text, Y+10 wp h100 Center, Transmit MORTARA HOLTER`nto Preventice
	Gui, Font, Normal
	
	;~ GuiControlGet, btn1, Pos, BodyGuardian
	GuiControlGet, btn2, Pos, MORTARA
	
	btnW := 90
	btnH := 50
	
	Gui, Add, Picture
		, % "Y" btn2Y+36 " X" btn2X+60 
		. " w" btnW " h" btnH " "
		. " +0x1000 vUpload gPhaseTask"
		, .\files\H3.png
	
	GuiControl 
	, % (wksloc="Main Campus" ? "Enable" : "Disable")  
	, Grab Preventice enrollments 
	
	tmpsite := RegExReplace(sites,"TRI\|")
	tmpsite := wksloc="Main Campus" ? tmpsite : RegExReplace(tmpsite,site.tab "\|",site.tab "||")
	Gui, Add, Tab3																		; add Tab bar with tracked sites
		, -Wrap x10 y10 w640 h400 vWQtab +HwndWQtab
		, % "ORDERS|" 
		. (wksloc="Main Campus" ? "INBOX||" : "") 
		. "ALL|" tmpsite
	GuiControlGet, wqDim, Pos, WQtab
	lvDim := "W" wqDimW-25 " H" wqDimH-35
	
	if (wksloc="Main Campus") {
		Gui, Tab, INBOX
		Gui, Add, Listview
			, % "-Multi Grid BackgroundSilver " lvDim " greadWQlv vWQlv_in hwndHLV_in"
			, filename|Name|MRN|DOB|Location|Study Date|wqid|Type|Need FTP
		Gui, ListView, WQlv_in
		LV_ModifyCol(1,"0")																; filename and path, "0" = hidden
		LV_ModifyCol(2,"160")															; name
		LV_ModifyCol(3,"60")															; mrn
		LV_ModifyCol(4,"80")															; dob
		LV_ModifyCol(5,"80")															; site
		LV_ModifyCol(6,"80")															; date
		LV_ModifyCol(7,"2")																; wqid
		LV_ModifyCol(8,"40")															; ftype
		LV_ModifyCol(9,"70 Center")														; ftp
	}
	
	Gui, Tab, ORDERS
	Gui, Add, Listview
		, % "-Multi Grid BackgroundSilver ColorRed " lvDim " greadWQorder vWQlv_orders hwndHLV_orders"
		, filename|Order Date|Name|MRN|Ordering Provider|Monitor
	Gui, ListView, WQlv_orders
	LV_ModifyCol(1,"0")																	; filename and path (hidden)
	LV_ModifyCol(2,"80")																; date
	LV_ModifyCol(3,"140")																; Name
	LV_ModifyCol(4,"60")																; MRN
	LV_ModifyCol(5,"100")																; Prov
	LV_ModifyCol(6,"70")																; Type
	
	Gui, Tab, ALL
	Gui, Add, Listview
		, % "-Multi Grid BackgroundSilver " lvDim " gWQtask vWQlv_all hwndHLV_all"
		, ID|Enrolled|FedEx|Uploaded|Notes|MRN|Enrolled Name|Device|Provider|Site
	Gui, ListView, WQlv_all
	LV_ModifyCol(1,"0")																	; wqid (hidden)
	LV_ModifyCol(2,"60")																; date
	LV_ModifyCol(3,"40 Center")															; FedEx
	LV_ModifyCol(4,"60")																; uploaded
	LV_ModifyCol(5,"40 Center")															; Notes
	LV_ModifyCol(6,"60")																; MRN
	LV_ModifyCol(7,"140")																; Name
	LV_ModifyCol(8,"130")																; Ser Num
	LV_ModifyCol(9,"100")																; Prov
	LV_ModifyCol(10,"80")																; Site
	CLV_all := new LV_Colors(HLV_all,true,false)
	CLV_all.Critical := 100
	
	Loop, parse, sites, |
	{
		i := A_index
		site := A_LoopField
		Gui, Tab, % site
		Gui, Add, Listview
			, % "-Multi Grid BackgroundSilver " lvDim " gWQtask vWQlv"i " hwndHLV"i
			, ID|Enrolled|FedEx|Uploaded|Notes|MRN|Enrolled Name|Device|Provider
		Gui, ListView, WQlv%i%
		LV_ModifyCol(1,"0")																	; wqid (hidden)
		LV_ModifyCol(2,"60")																; date
		LV_ModifyCol(3,"40 Center")															; FedEx
		LV_ModifyCol(4,"60")																; uploaded
		LV_ModifyCol(5,"40 Center")															; Notes
		LV_ModifyCol(6,"60")																; MRN
		LV_ModifyCol(7,"140")																; Name
		LV_ModifyCol(8,"130")																; Ser Num
		LV_ModifyCol(9,"100")																; Prov
		CLV_%i% := new LV_Colors(HLV%i%,true,false)
		CLV_%i%.Critical := 100
	}
	WQlist()
	
	Menu, menuSys, Add, Clean tempfiles, CleanTempFiles
	Menu, menuSys, Add, Change clinic location, changeLoc
	Menu, menuSys, Add, Generate late returns report, lateReport
	Menu, menuSys, Add, Generate registration locations report, regReport
	Menu, menuSys, Add, Update call schedules, updateCall
	Menu, menuHelp, Add, About TRRIQ, menuTrriq
	Menu, menuHelp, Add, Instructions..., menuInstr
		
	Menu, menuBar, Add, System, :menuSys
	Menu, menuBar, Add, Help, :menuHelp
	
	Gui, Menu, menuBar
	Gui, Show,, TRRIQ Dashboard
	
	SetTimer, idleTimer, 500
	return
}

PhaseGUIclose:
{
	MsgBox, 262161, Exit, Really quit TRRIQ?
	IfMsgBox, OK
	{
		checkPreventiceOrdersOut()
		cleanDone()
		eventlog("<<<<< Session end.")
		ExitApp
	}
	return
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
	ens:=wq.selectNodes("/root/pending/enroll")
	num := ens.length
	Loop, % num
	{
		Progress,,,% A_index "/" num 
		k := ens.item(A_Index-1)
		id	:= k.getAttribute("id")
		e := readWQ(id)
		dt := A_now
		dt -= e.date, Days
		if (instr(e.dev,"BG") && (dt > 45)) || (instr(e.dev,"Mortara") && (dt > 14))  {
			str .= e.site ",""" e.prov """," e.date ",""" e.name """," e.mrn "," e.dev "`n"
		}
	}
	progress, off
	tmp := path.holterPDF "late-" A_now ".csv"
	FileAppend, %str%, %tmp%
	eventlog("Generated missing devices report.")
	MsgBox, 262208, Missing devices report, Report saved to:`n%tmp%
	return
}

regReport:
{
	str := ""
	ens:=wq.selectNodes("//enroll")
	num := ens.length
	loop, % num
	{
		Progress,,,% A_index "/" num 
		k := ens.item(A_Index-1)
		id	:= k.getAttribute("id")
		e := readWQ(id)
		str .= e.site "," e.date "," "" e.prov "" "," e.dev "`n"
	}
	progress, off
	tmp := path.holterPDF "reg-" A_Now ".csv"
	FileAppend, %str%, %tmp%
	eventlog("Generated registrations report.")
	MsgBox, 262208, Registrations report, Report saved to:`n%tmp%
	return
}

PhaseTask:
{
	phase := A_GuiControl
	Gui, phase:Hide
	return
}

PhaseOrder:
{
	GuiControl, phase:Choose, WQtab, ORDERS
	return
}

PhaseRefresh:
{
	GuiControl, Text, ORDERS, No active orders
	GuiControl, Disable, orders
	WQlist()
}

idleTimer() {
	checkWQfile()
	x:=checkMUwin()
	;~ progress,,,% x
	;~ sleep 50
	;~ progress, off
	return
}

checkWQfile() {
	global wqfileDT
	FileGetTime, tmpdt, .\files\wqupdate												; get mod dt for "wqupdate"
	if (tmpdt > wqfileDT) {																; file is more recent than internal var
		wqfileDT := tmpdt																; set var to this date
		WQlist()																		; refresh list
	}
	return
}

setwqupdate() {
	global wqfileDT
	FileDelete, .\files\wqupdate
	FileAppend,,.\files\wqupdate
	wqfileDT := A_now
	return
}

checkMUwin() {
	global muwin
	static wintxt, tabtxt
	t0 := A_TickCount
	ui := MorUIgrab()																	; returns .tab, .txt, .TRct, .PRct
	
	if (ui.vis = wintxt) {																; form text unchanged
		t1 := A_TickCount-t0
		return t1
	}
	wintxt := ui.vis																	; reset text for wintxt comparison
	if !instr(ui.vis,"Second ID") {														; not on a form tab
		t1 := A_TickCount-t0
		return t1
	}
	RegExMatch(wintxt,"i)(Transfer|Prepare)",match)										; first string that matches will be in "match1"
	Gui, phase:Hide
	MortaraUpload(match1)
	
	return 
}

checkMachine() {
/*	Check if current machine has H3 software installed
	local machine names begin with EWCSS and Citrix machines start with PPWC,VMWIN10
*/
	global has_H3, wksVoid
	is_VM := ObjHasValue(wksVoid,A_ComputerName,1)

	if (A_UserName="tchun1") {
		; return
	}
	if (has_H3=false) {
		MsgBox 0x40030
			, Environment Error, % ""
			. (is_VM ? "Mortara H3 software not available on VDI/Citrix." : "Mortara H3 software not found!")
			. "`n`n"
			. "Switch to another computer if you will need to register/upload Mortara 24-hour Holter."
	}
	if (A_ComputerName~="EWC") {														; running on a local machine
		return																			; return successfully
	}
	else if (is_VM=true) {
		return
	}
	else {
		eventlog("Unique machine name.")
		return
	}
}

checkVersion(ver) {
	FileGetTime, chk, % A_ScriptName
	if (chk != ver) {
		MsgBox, 262193, New version!, There is an updated version of the script. `nRestart to launch new version?
		IfMsgBox, Ok
			run, % A_ScriptName
		ExitApp
	}
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
	
	;~ Gui, phase:Hide
	pt := readWQ(idx)
	idstr := "/root/pending/enroll[@id='" idx "']"
	
	list :=
	Loop, % (notes:=wq.selectNodes(idstr "/notes/note")).length 
	{
		k := notes.item(A_index-1)
		dt := parsedate(k.getAttribute("date"))
		list .= dt.mm "/" dt.dd ":" k.getAttribute("user") ": " k.text "`n"
	}

	choice := cmsgbox(pt.Name " " pt.MRN
			,	"Date: " niceDate(pt.date) "`n"
			.	"Provider: " pt.prov "`n"
			.	strQ(pt.FedEx,"  FedEx: ###`n")
			.   strQ(list,"Notes: ========================`n###`n")
			, "View/Add NOTE|"
			. "Log UPLOAD to Preventice|"
			. "Move to DONE list"
			, "Q")
	if (choice="xClose") {
		return
	}
	if instr(choice,"upload") {
		inputbox(inDT,"Upload log","`n`nEnter date uploaded to Preventice`n",niceDate(A_now))
		if (ErrorLevel) {
			return
		}
		wq := new XML("worklist.xml")
		if !IsObject(wq.selectSingleNode(idstr "/sent")) {
			wq.addElement("sent",idstr)
		}
		wq.setText(idstr "/sent",parseDate(inDT).YMD)
		wq.setAtt(idstr "/sent",{user:user})
		writeout(idstr,"sent")
		eventlog(pt.MRN " " pt.Name " study " pt.Date " uploaded to Preventice.")
		MsgBox, 4160, Logged, % pt.Name "`nUpload date logged!"
		setwqupdate()
		WQlist()
		return
	}
	if instr(choice,"note") {
		inputbox(note,"Communication note"
			, strQ(list,"###====================================`n") "`nEnter a brief communication note:`n","")
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
		setwqupdate()
		WQlist()
		return
	}
	if instr(choice,"done") {
		reason := cmsgbox("Reason"
				, "What is the reason to remove this record from the active worklist?"
				, "Report in Epic|"
				. "Device missing|"
				. "Other (explain)"
				, "E")
		if (reason="xClose") {
			return
		}
		if instr(reason,"Other") {
			reason:=""
			inputbox(reason,"Clear record from worklist","Enter the reason for moving this record","")
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
		setwqupdate()
		WQlist()
	}
return	
}

WQlist() {
	global
	local k, ens, e0, id, now, dt, site, fnID, res, key, val, full, wqfiles, lvDim
		, late_BGH := 45
		, late_BGM := 21
		, late_Mortara := 14
	wqfiles := []
	fldval := {}
	GuiControlGet, wqDim, Pos, WQtab
	lvDim := "W" wqDimW-25 " H" wqDimH-35
	
	checkversion(runningVer)															; make sure we are running latest version

	Progress,,% " ",Scanning worklist...
	
	fileCheck()
	FileOpen(".lock", "W")																; Create lock file.
	
	wq := new XML("worklist.xml")														; refresh WQ
	
	readPrevTxt()																		; read prev.txt from website
	
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
	
	/*	Scan Incoming ORDERS from EPIC
	*/
	Gui, ListView, WQlv_orders
	LV_Delete()
	
	if !IsObject(wq.selectSingleNode("/root/orders")) {
		wq.addElement("orders","/root")
	}
	
	Loop, files, % path.EpicHL7in "*"													; First pass: process new files
	{
		e0 := {}
		fileIn := A_LoopFileName
		if RegExMatch(fileIn,"_([a-zA-Z0-9]{4,})Z.hl7",i) {								; skip old files
			continue
		}
		processhl7(A_LoopFileFullPath)
		e0:=parseORM()
		if InStr(sites0, e0.loc) {														; skip non-tracked orders
			FileMove, %A_LoopFileFullPath%, % ".\tempfiles\" e0.mrn "_" e0.nameL "_" A_LoopFileName, 1
			eventlog("Non-tracked order " fileIn " moved to tempfiles. " e0.loc " " e0.mrn " " e0.nameL)
			continue
		}
		eventlog("New order " fileIn ". " e0.name " " e0.mrn )
		
		loop, % (ens:=wq.selectNodes("/root/pending/enroll")).Length					; find enroll nodes with result but no order
		{
			k := ens.item(A_Index-1)
			if IsObject(k.selectSingleNode("accession")) {								; skip nodes that already have accession
				continue
			}
			e0.match_NM := fuzzysearch(e0.name,format("{:U}",k.selectSingleNode("name").text))
			e0.match_MRN := fuzzysearch(e0.mrn,k.selectSingleNode("mrn").text)
			if (e0.match_NM > 0.15) && (e0.match_MRN > 0.15) {							; Name and MRN each vary by more than 15%
				continue
			}
			dt0 := k.selectSingleNode("date").text
			dt0 -= e0.date, days
			if abs(dt0) > 5 {															; Date differs by more than 5d
				Continue
			}

			id := k.getAttribute("id")
			e0.match_UID := true
			
			wqSetVal(id,"order",e0.order)
			wqSetVal(id,"accession",e0.accession)
			wqSetVal(id,"acct",e0.acct)
			wqSetVal(id,"acctnum",e0.accountnum)
			wqSetVal(id,"encnum",e0.encnum)
			k.setAttribute("id",e0.UID)
			eventlog("Found pending/enroll=" id " that matches new Epic order " e0.order ". " e0.match_NM)
			eventlog("enroll id " id " changed to " e0.UID)
			break
		}
		if (e0.match_UID) {
			FileMove, %A_LoopFileFullPath%, .\tempfiles\*, 1
			eventlog("Moved: " A_LoopFileFullPath)
			continue
		}
		
		e0.orderNode := "/root/orders/enroll[order='" e0.order "']"
		if IsObject(k:=wq.selectSingleNode(e0.orderNode)) {								; ordernum node exists
			e0.nodeCtrlID := k.selectSingleNode("ctrlID").text
			if (e0.CtrlID < e0.nodeCtrlID) {											; order CtrlID is older than existing, somehow
				FileDelete, % path.EpicHL7in fileIn
				eventlog("Order msg " fileIn " is outdated. " e0.name)
				continue
			}
			if (e0.orderCtrl="CA") {													; CAncel an order
				FileDelete, % path.EpicHL7in fileIn										; delete this order message
				FileDelete, % path.EpicHL7in "*_" e0.UID "Z.hl7"						; and the previously processed hl7 file
				removeNode(e0.orderNode)												; and the accompanying node
				eventlog("Cancelled order " e0.order ". " e0.name)
				continue
			}
			FileDelete, % path.EpicHL7in "*_" e0.UID "Z.hl7"							; delete previously processed hl7 file
			removeNode(e0.orderNode)													; and the accompanying node
			eventlog("Cleared order " e0.order " node. " e0.name)
		}
		if (e0.orderCtrl="XO") {														; change an order
			e0.orderNode := "/root/orders/enroll[accession='" e0.accession "']"
			k := wq.selectSingleNode(e0.orderNode)
			e0.nodeUID := k.getAttribute("id")
			FileDelete, % path.EpicHL7in "*_" e0.nodeUID "Z.hl7"
			removeNode(e0.orderNode)
			eventlog("Removed node id " e0.nodeUID " for replacement. " e0.name)
		}
		
		newID := "/root/orders/enroll[@id='" e0.UID "']"								; otherwise create a new node
			wq.addElement("enroll","/root/orders",{id:e0.UID})
			wq.addElement("order",newID,e0.order)
			wq.addElement("accession",newID,e0.accession)
			wq.addElement("ctrlID",newID,e0.CtrlID)
			wq.addElement("date",newID,e0.date)
			wq.addElement("name",newID,e0.name)
			wq.addElement("mrn",newID,e0.mrn)
			wq.addElement("sex",newID,e0.sex)
			wq.addElement("dob",newID,e0.dob)
			wq.addElement("mon",newID,e0.mon)
			wq.addElement("prov",newID,e0.prov)
			wq.addElement("provname",newID,e0.provname)
			wq.addElement("site",newID,e0.loc)
			wq.addElement("acct",newID,e0.acct)
			wq.addElement("acctnum",newID,e0.accountnum)
			wq.addElement("encnum",newID,e0.encnum)
			wq.addElement("ind",newID,e0.ind)
		eventlog("Added order ID " e0.UID ". " e0.name)
		
		fileOut := (e0.mon="CUTOVER" ? "done\" : "")
			. e0.MRN "_" 
			. fldval["PID_nameL"] "^" fldval["PID_nameF"] "_"
			. e0.date "_"
			. e0.uid 																	; new ORM filename ends with _[UID]Z.hl7
			. "Z.hl7"
		
		FileMove, %A_LoopFileFullPath%													; and rename ORM file
			, % path.EpicHL7in . fileOut
		
	}
	
	loop, Files, % path.EpicHL7in "*Z.hl7"												; Second pass: scan *Z.hl7 files
	{
		e0 := {}
		fileIn := A_LoopFileName
		if RegExMatch(fileIn,"_([a-zA-Z0-9]{4,})Z.hl7",i) {								; file appears to have been parsed
			e0 := readWQ(i1)
		} else {
			continue
		}
		
		if instr(sites0,e0.site) {														; sites0 location
			FileMove, %A_LoopFileFullPath%, .\tempfiles, 1
			removeNode("/root/orders/enroll[@id='" i1 "']")
			eventlog("Non-tracked order " fileIn " moved to tempfiles.")
			continue
		}
		if (e0.node ~= "pending|done") {												; remnant orders file
			FileMove, %A_LoopFileFullPath%, .\tempfiles, 1
			eventlog("Leftover HL7 file " fileIn " moved to tempfiles.")
			continue
		}
		
		LV_Add(""
			, path.EpicHL7in . fileIn													; filename and path to HolterDir
			, e0.date																	; date
			, e0.name																	; name
			, e0.mrn																	; mrn
			, e0.provname																; prov
			, (e0.mon="HOL" ? "24-hr "													; monitor type
				: e0.mon="BGH" ? "30-day "
				: e0.mon~="BGM|ZIO" ? "14-day "
				: "")
				. e0.mon
			, "")																		; fulldisc present, make blank
		GuiControl, Enable, Register
		GuiControl, Text, Register, Go to ORDERS tab
	}
	
	loop, % (ens:=wq.selectNodes("/root/orders/enroll")).Length							; Third pass: remove extraneous orders
	{
		e0 := {}
		k := ens.item(A_Index-1)
		e0.uid := k.getAttribute("id")
		e0.order := k.selectSingleNode("order").text
		e0.accession := k.selectSingleNode("accession").text
		e0.name := k.selectSingleNode("name").text
		
		if IsObject(wq.selectSingleNode("/root/pending/enroll[order='" e0.order "'][accession='" e0.accession "']")) {
			eventlog("Order node " e0.uid " " e0.name " already found in pending.")
			removenode("/root/orders/enroll[@id='" e0.uid "']")
		}
		if IsObject(wq.selectSingleNode("/root/done/enroll[order='" e0.order "'][accession='" e0.accession "']")) {
			eventlog("Order node " e0.uid " " e0.name " already found in done.")
			removenode("/root/orders/enroll[@id='" e0.uid "']")
		}
	}
	
	WriteSave(wq)
	FileDelete, .lock
	
	checkPreventiceOrdersOut()															; check registrations that failed upload to Preventice
	
	if (wksloc="Main Campus") {
		
	Gui, ListView, WQlv_in
	LV_Delete()																			; clear the INBOX entries
	
/*	Process each incoming .hl7 RESULT from PREVENTICE
*/
	loop, Files, % path.PrevHL7in "*.hl7"
	{
		fileIn := A_LoopFileName
		x := StrSplit(fileIn,"_")
		if !(id := hl7dirMap[fileIn]) {													; will be true if have found this wqid in this instance, else null
			fileread, tmptxt, % path.PrevHL7in fileIn
			obr:= strsplit(stregX(tmptxt,"\R+OBR",1,0,"\R+",0),"|")						; get OBR segment
			obr_req := trim(obr.3," ^")													; wqid from Preventice registration (PV1_19)
			obr_prov := strX(obr.17,"^",1,1,"^",1)
			obr_site := strX(obr_prov,"-",0,1,"",0)
			pv1:= strsplit(stregX(tmptxt,"\R+PV1",1,0,"\R+",0),"|")						; get PV1 segment
			pv1_dt := SubStr(pv1.40,1,8)												; pull out date of entry/registration (will not match for send out)
			
			if instr(sites0,obr_site) {
				eventlog("Unregistered Sites0 report (" fileIn " - " obr_site ")")
				FileMove, % path.PrevHL7in fileIn, .\tempfiles\%fileIn%, 1
				continue
			}
			if (readWQ(obr_req).mrn) {													; check if obr_req is valid wqid
				id := obr_req
				hl7dirMap[fileIn] := id
			} 
			else if (id := findWQid(pv1_dt,x.3).id) { 									; try to find wqid based on date in PV1.40 and mrn
				hl7dirMap[fileIn] := id
			}
			else {																		; can't find wqid, just admit defeat
				id :=
			}
		}
		res := readWQ(id)																; wqid should always be present in hl7 downloads
		if (res.node="done") {															; skip if DONE, might be currently in process 
			eventlog("Report already done (" id ": " res.name " - " res.mrn ", " res.date ")")
			eventlog("WQlist removing " fileIn)
			FileMove, % path.PrevHL7in fileIn, .\tempfiles\%fileIn%, 1
			continue
		}
		FileGetSize,full,% path.PrevHL7in fileIn,M
		
		LV_Add(""
			, path.PrevHL7in fileIn														; path and filename
			, strQ(res.Name,"###", x.1 ", " x.2)										; last, first
			, strQ(res.mrn,"###",x.3)													; mrn
			, strQ(niceDate(res.dob),"###",niceDate(x.4))								; dob
			, strQ(res.site,"###",obr_site)												; site
			, strQ(niceDate(res.date),"###",niceDate(SubStr(x.5,1,8)))					; study date
			, id																		; wqid
			, (res.dev~="Heart|IMD POST EVENT") ? "BGH"									; extracted
			: (res.dev~="Mortara") ? "HOL"
			: (res.dev~="Mini") ? "MINI"
			: "HL7"
			, (res.dev~="Mortara")&&(full<3) ? "X":"")									; flag FTP if Mortara but filesize <3 Meg
		wqfiles.push(id)
	}
	
/*	Scan Holter PDFs folder for additional files
*/
	findfullPDF()																		; read Holter PDF dir into pdfList
	for key,val in pdfList
	{
		RegExMatch(val,"O)_WQ([A-Z0-9]+)_([A-Z])\.pdf",fnID)							; get filename WQID if PDF has already been renamed (fnid.1 = wqid, fnid.2 = type)
		id := fnID.1
		ftype := (fnID.2="H") ? "HOL"													; type of file based on fnID label
				: (fnID.2="Z") ? "ZIO"
				: (fnID.2="E") ? "BGH"
				: (fnID.2="M") ? "MINI"
				: "???"																	; could condense as ftype := {"H":"PDF","Z":"ZIO","E":"CEM","M":"MINI"}[fnID.2]
		if (k:=ObjHasValue(wqfiles,id)) {												; found a PDF file whose wqid matches an hl7 in wqfiles
			LV_Modify(k,"Col9","")														; clear the "X" in the FullDisc column
			continue																	; skip rest of processing
		}
		res := readwq(id)																; get values for wqid if valid, else null
		
		LV_Add(""
			, path.holterPDF val														; filename and path to HolterDir
			, strQ(res.Name,"###",strX(val,"",1,0,"_",1))								; name from wqid or filename
			, strQ(res.mrn,"###",strX(val,"_",1,1,"_",1))								; mrn
			, strQ(res.dob,"###")														; dob
			, strQ(res.site,"###","???")												; site
			, strQ(nicedate(res.date),"###")											; study date
			, id																		; wqid
			, ftype																		; study type
			, "")																		; fulldisc present, make blank
		if (id) {
			wqfiles.push(id)															; add non-null wqid to wqfiles
		}
	}
	LV_ModifyCol(6,"Sort")																; date

	}	; <-- finish Main Campus Inbox
	
/*	Now scan <pending/enroll> nodes
*/
	Gui, ListView, WQlv_all
	LV_Delete()
	
	Loop, parse, sites, |
	{
		i := A_index
		site := A_LoopField
		Gui, ListView, WQlv%i%
		LV_Delete()																		; refresh each respective LV
		Loop, % (ens:=wq.selectNodes("/root/pending/enroll[site='" site "']")).length
		{
			k := ens.item(A_Index-1)
			id	:= k.getAttribute("id")
			e0 := readWQ(id)
			dt := A_now
			dt -= e0.date, Days
			e0.dev := RegExReplace(e0.dev,"BodyGuardian","BG")
			;~ if (instr(e0.dev,"BG") && (dt < 30)) {										; skip BGH less than 30 days
				;~ continue
			;~ }
			CLV_col := ""
			if (instr(e0.dev,"Heart") && (dt > late_BGH))
			|| (instr(e0.dev,"Mortara") && (dt > late_Mortara)) 
			|| (instr(e0.dev,"Mini") && (dt > late_BGM)) {
				CLV_col := "red"
			}
			
			Gui, ListView, WQlv%i%
			LV_Add(""
				,id
				,e0.date
				,strQ(e0.fedex,"X")
				,e0.sent
				,strQ(e0.notes,"X")
				,e0.mrn
				,e0.name
				,e0.dev
				,e0.prov
				,e0.site)
			if (CLV_col) {
				CLV_%i%.Row(LV_GetCount(),,CLV_col)
			}
			Gui, ListView, WQlv_all														
			LV_Add(""
				,id
				,e0.date
				,strQ(e0.fedex,"X")
				,e0.sent
				,strQ(e0.notes,"X")
				,e0.mrn
				,e0.name
				,e0.dev
				,e0.prov
				,e0.site)
			if (CLV_col) {
				CLV_all.Row(LV_GetCount(),,CLV_col)
			}
		}
		Gui, ListView, WQlv%i%
		LV_ModifyCol(2,"Sort")
	}
	Gui, ListView, WQlv_all														
	LV_ModifyCol(2,"Sort")
	
	tmp := parsedate(wq.selectSingleNode("/root/pending").getAttribute("update"))
	GuiControl, Text, PhaseNumbers
		,	% "Patients registered in Preventice (" wq.selectNodes("/root/pending/enroll").length ")`n"
		.	"Preventice update: " tmp.mm "/" tmp.dd " @ " tmp.time "`n"
	
	fileIn :=
	progress, off
	return
}

checkPreventiceOrdersOut() {
	global path
	
	loop, files, % path.PrevHL7out "Failed\*.txt"
	{
		filenm := A_LoopFileName
		filenmfull := A_LoopFileFullPath
		eventlog("Resending failed registration: " filenm)
		FileMove, % filenmfull, % path.PrevHL7out filenm
	}
	
	return
}

cleanDone() {
	global wq
	
	fileCheck()
	FileOpen(".lock", "W")																; Create lock file.
	
	if fileexist("archive.xml") {
		arc := new XML("archive.xml")
	} else {
		arc := new XML("<root/>")
		arc.addElement("done","/root")
		arc.save("archive.xml")
	}
	
	wq := new XML("worklist.xml")														; get most recently saved
	ens := wq.selectNodes("/root/done/enroll")
	t := ens.length
	progress,,% " ",Cleaning old records
	loop, % t
	{
		progress, % (A_Index/t)*100
		en := ens.item(A_Index-1)
		dt := en.selectSingleNode("date").text
		name := en.selectSingleNode("name").text
		dtDiff := A_now
		dtDiff -= dt, Days
		if (dtDiff<180) {																; skip dates less than 180 days
			continue
		}
		clone := en.cloneNode(true)
		arc.selectSingleNode("/root/done").appendChild(clone)
		en.parentNode.removeChild(en)
		eventlog("Removed old record (" dtDiff " days) for " name " " dt ".")
	}
	
	arc.save("archive.xml")
	writeSave(wq)
	wq := new XML("worklist.xml")
	FileDelete, .lock
	
	return
}


readPrevTxt() {
	global wq
	
	filenm := ".\files\prev.txt"
	prevtxtdt := wq.selectSingleNode("/root/pending").getAttribute("update")
	FileGetTime, filedt, % filenm
	if (filedt=prevtxtdt) {																; update matches filedt means no change
		return
	}
	
	Progress,,% " ",Updating Preventice data...
	FileRead, txt, % filenm
	StringReplace txt, txt, `n, `n, All UseErrorLevel
	n := ErrorLevel
	
	k := wq.selectSingleNode("/root/inventory")											; create fresh inventory node
	k.parentNode.removeChild(k)
	wq.addElement("inventory","/root")
	
	loop, read, % ".\files\prev.txt"
	{
		Progress, % 100*A_Index/n
		
		k := A_LoopReadLine
		if (k~="^enroll\|") {
			parsePrevEnroll(k)
		}
		else if (k~="^dev\|") {
			parsePrevDev(k)
		}
	}
	
	loop, % (devs := wq.selectNodes("/root/inventory/dev")).length						; Find dev that already exist in Pending
	{
		k := devs.item(A_Index-1)
		dev := k.getAttribute("model")
		ser := k.getAttribute("ser")
		if IsObject(wq.selectSingleNode("/root/pending/enroll[dev='" dev " - " ser "']")) {	; exists in Pending
			k.parentNode.removeChild(k)
			eventlog("Removed inventory ser " ser)
		}
	}
	wq.selectSingleNode("/root/pending").setAttribute("update",filedt)					; set pending[@update] attr
	wq.selectSingleNode("/root/inventory").setAttribute("update",filedt)				; set pending[@update] attr
	
return	
}

parsePrevEnroll(txt) {
	global wq
	el := StrSplit(txt,"|")
	res := {  date:parseDate(el.2).YMD
			, name:parsename(el.3).lastfirst
			, mrn:el.4
			, dev:el.5
			, prov:filterProv(el.6).name
			, site:filterProv(el.6).site }
	
	if (res.dev~="-$") {																; e.g. "Body Guardian Mini -"
		res.dev .= res.name																; append string so will not match in enrollcheck
	}
	
	/*	Check whether any params match this device
	*/
		if enrollcheck("[name='" res.name "']"											; 6/6 perfect match
			. "[mrn='" res.mrn "']"
			. "[date='" res.date "']"
			. "[dev='" res.dev "']"
			. "[prov='" res.prov "']"
			. "[site='" res.site "']" ) {
			return
		}
		if (id:=enrollcheck("[name='" res.name "']"										; 4/6 perfect match
			. "[mrn='" res.mrn "']"														; everything but PROV or SITE
			. "[date='" res.date "']"
			. "[dev='" res.dev "']" )) {
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			parsePrevElement(id,en,res,"prov")
			parsePrevElement(id,en,res,"site")
			return
		}
		if (id:=enrollcheck("[mrn='" res.mrn "']"										; Probably perfect MRN+S/N+DATE
			. "[date='" res.date "']"
			. "[dev='" res.dev "']" )) {
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			parsePrevElement(id,en,res,"name")
			parsePrevElement(id,en,res,"prov")
			parsePrevElement(id,en,res,"site")
			return
		}
		if (id:=enrollcheck("[mrn='" res.mrn "'][date='" res.date "']")) {				; MRN+DATE, no S/N
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			if (en.node="orders") {														; falls through if not in <pending> or <done>
				addPrevEnroll(id,res)													; create a <pending> record
				wqSetVal(id,"acct",en.acct)
				wqSetVal(id,"order",en.order)
				wqSetVal(id,"accession",en.accession)
				wqSetVal(id,"accountnum",en.acctnum)
				wqSetVal(id,"encnum",en.encnum)
				removeNode("/root/orders/enroll[@id='" id "']")
				eventlog("Moved Order ID " id " for " en.name " to Pending.")
			}
			parsePrevElement(id,en,res,"dev")
			return
		}
		if (id:=enrollcheck("[date='" res.date "'][dev='" res.dev "']")) {				; DATE+S/N, no MRN
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			parsePrevElement(id,en,res,"mrn")
			return
		} 
		if (id:=enrollcheck("[mrn='" res.mrn "'][dev='" res.dev "']")) {				; MRN+S/N, no DATE match
			en:=readWQ(id)
			if (en.node="done") {
				return
			}
			dt0:=res.date
			dt0 -= en.date, days
			if abs(dt0) < 5 {															; res.date less than 5d from en.date
				parsePrevElement(id,en,res,"date")										; prob just needs a date adjustment
				return
			}
		}
		if (id:=wq.selectSingleNode("/root/orders/enroll[mrn='" res.mrn "']").getAttribute("id")) {
			en:=readWQ(id)																; MRN found in Orders
			dt0:=res.date
			dt0 -= en.date, days
			
			if abs(dt0) < 5 {															; res.date less than 5d from en.date
				addPrevEnroll(id,res)													; create a <pending> record
				wqSetVal(id,"acct",en.acct)
				wqSetVal(id,"order",en.order)
				wqSetVal(id,"accession",en.accession)
				wqSetVal(id,"accountnum",en.acctnum)
				wqSetVal(id,"encnum",en.encnum)
				wqSetVal(id,"prov",en.provname)
				wqSetVal(id,"dev",res.dev)
				wqSetVal(id,"date",res.date)
				wqSetVal(id,"ind",en.ind)
				removeNode("/root/orders/enroll[@id='" id "']")
				eventlog("Order ID " id " for " en.name " " en.mrn " matched MRN only, moved to Pending.")
				return
			}
		}																				; anything else is probably a new registration
		
	/*	No match (i.e. unique record)
	 *	add new record to PENDING
	 */
		id := makeUID()
		addPrevEnroll(id,res)
	
	return
}

addPrevEnroll(id,res) {
	global wq
	
	newID := "/root/pending/enroll[@id='" id "']"
	wq.addElement("enroll","/root/pending",{id:id})
	wq.addElement("date",newID,res.date)
	wq.addElement("name",newID,res.name)
	wq.addElement("mrn",newID,res.mrn)
	wq.addElement("dev",newID,res.dev)
	wq.addElement("prov",newID,res.prov)
	wq.addElement("site",newID,res.site)
	wq.addElement("webgrab",newID,A_now)
	
	eventlog("Found new web registration " res.mrn " " res.name " " res.date ".")
	return
}

parsePrevElement(id,en,res,el) {
	global wq
	
	if (res[el]=en[el]) {																; Attr[el] is same in EN (wq) as RES (txt)
		return																			; don't do anything
	}
	
	wqSetVal(id,el,res[el])
	eventlog(en.name " (" id ") changed WQ " el " '" en[el] "' ==> '" res[el] "'")
	
	return
}

parsePrevDev(txt) {
	global wq
	el := StrSplit(txt,"|")
	dev := el.2
	ser := el.3
	res := dev " - " ser

	if IsObject(wq.selectSingleNode("/root/inventory/dev[@ser='" ser "']")) {			; already exists in Inventory
		return
	}
	
	wq.addElement("dev","/root/inventory",{model:dev,ser:ser})
	;~ eventlog("Added new Inventory dev " ser)
	
	return
}

makeUID() {
	global wq
	
	Loop
	{
		Random, num1, 10000, 99999
		Random, num2, 10000, 99999
		Random, num3, 10000, 99999
		num := num1 . num2 . num3
		id := toBase(num,36)
		if IsObject(wq.selectSingleNode("//enroll[id='" id "']")) {
			eventlog("UID " id " already in use.")
			continue
		} 
		else {
			break
		}
	}
	return id
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
	SplitPath,fileIn,fnam,,fExt,fileNam
	
	wq := new XML("worklist.xml")														; refresh WQ
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
	if (fldval.node = "done") {															; task has been done already by another user
		MsgBox, 262208, Completed, File has already been processed!
		WQlist()																		; refresh list and return
		return
	}
	
	if (fExt="hl7") {																	; hl7 file (could still be Holter or CEM)
		eventlog("===> " fnam )
		Gui, phase:Hide
		
		progress, 25 , % fnam, Extracting data
		processHL7(path.PrevHL7in . fnam)												; extract DDE to fldVal, and PDF into hl7Dir
		moveHL7dem()																	; prepopulate the fldval["dem-"] values
		
		checkEpicOrder()																; check for presence of valid Epic order
		
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
		makeORU(wqid)
		gosub outputfiles																; generate and save output CSV, rename and move PDFs
		
		if (fldval.oldUID) {
			MsgBox, 262192
				, Cutover study
				, % "Successfully processed Epic cutover report.`n`n"
				. "1) Return to Epic Tech Work List.`n"
				. "2) End Study for """ fldval["dem-name"] """.`n`n"
				. "3) Complete tech biller for """ strX(fldval.obr4,"^",1,1,"^",1,1) "`n"
		}
	}
	
	return
}

readWQorder() {
/*	Retrieve info from WQlist line
*/
	global wq, fldval, ptDem, sitesLong, mwuPhase
	fldval := {}
	ptDem := {}
	pt := {}
	
	
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
	SplitPath,fileIn,fnam,,fExt,fileNam
	
	Gui, phase:Destroy
	
	wq := new XML("worklist.xml")														; refresh WQ
	processhl7(fileIn)																	; read HL7 OBX into fldval
	ptDem:=parseORM()																	; read fldval into ptDem
	ptDem.filename := fileIn
	ptDem.Provider := ptDem.provname
	
	if (ptDem.monitor~="i)HOL") {														; for Mortara Holter
		eventlog("Start Mortara prepare.")
		mwuPhase := "Prepare"
		mortaraUpload(mwuPhase)
	} 
	else if (ptDem.monitor~="i)BGM") {													; for BG Mini (and maybe Zio)
		BGregister("BGM")
	}
	else if (ptDem.monitor~="i)BGH") {													; for BG Heart
		BGregister("BGH")
	}
	
	wqlist()
	return
}

checkEpicOrder() {
/*	Check for presence of valid <pending> node (has accession number)
	
	Check for <orders> node that matches the parsed ORU
	
	"In-flight" legacy results will not have existing Epic orders
	Epic order number necessary to move forward with resulting
	If needed, MA will place order and check-in study to create ORM
*/
	global fldval, wq
	
	if (fldval.accession) {																; Accession number exists, return to processing
		return
	}
	
	/*	Search for <orders/enroll> node that matches name in this result
		Only occurs if ORM parsed but has no matching registration
	*/
	loop, % (ens := wq.selectNodes("/root/orders/enroll")).Length
	{
		en := ens.item(A_Index-1)
		en_id := en.getAttribute("id")
		en_name := en.selectSingleNode("name").text
		en_date := en.selectSingleNode("date").text
		en_mrn := en.selectSingleNode("mrn").text
		en_mon := en.selectSingleNode("mon").text										; en_mon=order HOL|BGM|BGH 
		fld_mon := (en_mon~="HOL|BGM") ? "Holter"										; fld_mon => Holter|CEM
				:  (en_mon~="BGH") ? "CEM"  											; fldval.OBR_TestCode=Holter|CEM
				: ""
		
		if (en_name = fldval["dem-name"]) {
			eventlog("Found order for " en_name " (" en_id "), " en_mon ".")
			progress, hide
			MsgBox, 262196, 
			, % "Found this:`n"
			.   "   " en_name "`n"
			.   "   " parseDate(en_date).MDY "`n"
			.   "   " en_mon "`n`n"
			. "Use this order?"
			IfMsgBox, Yes
			{
				fldval.order := en.selectSingleNode("order").text
				fldval.accession := en.selectSingleNode("accession").text
				fldval.acct := fldval.site "_" fldval.order "-" fldval.accession
				wqsetval(fldval.wqid,"order",fldval.order)
				wqsetval(fldval.wqid,"accession",fldval.accession)
				wqsetval(fldval.wqid,"acct",fldval.acct)
				writeOut("/root/pending","enroll[@id='" fldval.wqid "']")
				eventlog("Used order.")
				return
			} else {
				eventlog("Cancelled.")
			}
			progress, show
		}
	}
	
	/*	Check if valid order already exists
		Tech must find Order Report that includes "Order #" and "Accession #"
		Return if found, or Cancel to move on
	*/
	Loop
	{
		SetTimer, checkEpicClip, 500
		progress, hide
		MsgBox, 262193
			, Check for Epic order
			, % "Check to see if patient has existing order.`n`n"
			. "1) Search for """ fldval["dem-name"] """.`n"
			. "2) Under Encounters, select the correct encounter on " parsedate(fldval.date).mdy ".`n"
			. "3) Click on the Holter/Event Monitor order in Orders Performed.`n"
			. "4) Right-click within the order, and select 'Copy all'.`n`n"
			. "Select [Cancel] if there is no existing order."
		SetTimer, checkEpicClip, off
		IfMsgBox, Cancel
		{
			break
		}
		if (fldval.accession) {
			return
		}
	}
	
	/*	Can't find an order, use Cutover order method
		This is the last resort, as it creates a lot of confusion with results
	*/
	progress, hide
	MsgBox, 262196, Missing EPIC order, Do you want to start CUTOVER process?
	IfMsgBox, Yes
	{
		MsgBox, 262192
			, Needs CUTOVER order
			, % "Study registered before Epic Go-Live. Valid Epic order required.`n`n"
			. "1) Open ""Anc Orders"" from Epic top toolbar.`n"
			. "2) Search for """ fldval["dem-name"] """.`n`n"
			. "3) Click the ""New Order"" button and fill out:`n"
			. "   - Referring Prov: """ fldval["dem-ordering"] """`n"
			. "   - Department: ordering clinic`n"
			. "   - Procedure: Cutover Holter/Event Monitor`n"
			. "   - Diagnosis: patient's primary diagnosis`n"
			. "   - Reason for exam: choose the indication`n"
			. "   - click ""Accept""`n`n"
			. "4) Fill in """ fldval["dem-ordering"] """ as the Authorizing Provider.`n"
			. "5) Click the ""Add-on"" button to move to the Technician Work List.`n`n"
			. "6) You may need to complete the ""Check-in"" for the patient.`n`n"
			. "7) Click [OK] here when you have done this."
		
		wqid := fldval.wqid
		wqsetval(wqid,"oldUID",wqid)
		writeOut("/root/pending","enroll[@id='" wqid "']")
		eventlog("Created CUTOVER order for " wqid)
		progress, show
		
		gosub MainLoop
		return
	} else {
		eventlog("Did not want CUTOVER order.")
		return
	}
	return
}

checkEpicClip() {
	global fldval
	
	i := substr(clipboard,1,350)
	if instr(i,"Order #") {
		settimer, checkEpicClip, off
		ControlClick, OK, Check for Epic order
		ordernum := trim(stregX(i,"Order #:",1,1,"Accession",1))
		accession := trim(stregX(i,"Accession #:",1,1,"\R+",1))
		RegExMatch(i,"im)^(.*)\R+Order #",dev)
		date := parsedate(stregX(i,"Ordered On ",1,1,"\s",1)).MDY
		mrn := trim(stregX(i,"MRN:",1,1,"\R+",1))
		name := stRegX(i,"^",1,0,"`r`nMRN:",1)
		clipboard :=
		
		MsgBox, 262180
			, Order found, % ""
			. "Type: " dev1 "`n"
			. "Date placed: " date "`n"
			. "Order #" ordernum "`n"
			. "Accession #" accession "`n`n"
			. "Use this order?"
		IfMsgBox, yes
		{
			fldval.order := ordernum
			fldval.accession := accession
			fldval.acct := fldval.site "_" fldval.order "-" fldval.accession
			wqsetval(fldval.wqid,"order",fldval.order)
			wqsetval(fldval.wqid,"accession",fldval.accession)
			wqsetval(fldval.wqid,"acct",fldval.acct)
			eventlog("Grabbed order #" fldval.order ", accession #" fldval.accession)

			if (name!=fldval.name) {
				MsgBox, 0x40031
					, Name Mismatch, % ""
					. "Correct the name`n     '" fldval["dem-Name"] "'`n"
					. "to this:`n     '" name "'"
				IfMsgBox, OK
				{
					fldval["dem-Name"] := name
					fldval["dem-NameL"] := ParseName(name).last
					fldval["dem-NameF"] := ParseName(name).first
					wqSetVal(fldval.wqid,"name",fldval["dem-Name"])								; make sure name matches Epic result
					eventlog("dem-Name changed '" fldval["dem-Name"] "' ==> '" name "'")
				}
			}
			writeOut("/root/pending","enroll[@id='" fldval.wqid "']")
		}
	}
	return
}

parseORM() {
/*	parse fldval values to values
	including aliases for both WQlist and readWQorder
*/
	global fldval, sitesLong, indCodes
	
	monType:=(tmp:=fldval.OBR_TestName)~="i)14 DAY" ? "BGM"
		: tmp~="i)15 DAY" ? "BGM"
		: tmp~="i)24 HOUR" ? "HOL"
		: tmp~="i)RECORDER|EVENT" ? "BGH"
		: tmp~="i)CUTOVER" ? "CUTOVER"
		: ""
	
	switch fldval.PV1_PtClass
	{
		case "O":
			encType := "Outpatient"
			location := sitesLong[fldval.PV1_Location]
		case "I":
			encType := "Inpatient"
			location := fldval.PV1_Location
		case "DS":
			encType := "Outpatient"
			location := "MAIN"
		case "E":
			encType := "Inpatient"
			location := "Emergency"
		default:
			encType := "Outpatient"
			location := fldval.PV1_Location
	}
	prov := strQ(fldval.ORC_ProvCode
			, fldval.ORC_ProvCode "^" fldval.ORC_ProvNameL "^" fldval.ORC_ProvNameF
			, fldval.OBR_ProviderCode "^" fldval.OBR_ProviderNameL "^" fldval.OBR_ProviderNameF)
	provname := strQ(fldval.ORC_ProvCode
			, fldval.ORC_ProvNameL strQ(fldval.ORC_ProvNameF, ", ###")
			, fldval.OBR_ProviderNameL strQ(fldval.OBR_ProviderNameF, ", ###"))
	;~ location := (encType="Outpatient") ? sitesLong[fldval.PV1_Location]
		;~ : encType
		
	if !(indication:=strQ(fldval.OBR_ReasonCode,"###") strQ(fldval.OBR_ReasonText,"^###")) {
		indText := objhasvalue(fldval,"^Reason for exam","RX")
		indText := RegExReplace(fldval[indText],"Reason for exam->")
		
		indCode := objhasvalue(indCodes,indText,"RX")
		indCode := strX(indCodes[indCode],"",1,0,":",1,1)
		
		indication := strQ(indCode,"###") strQ(indText,"^###")
	}
	
	return {date:parseDate(fldval.PV1_DateTime).YMD
		, encDate:parseDate(fldval.PV1_DateTime).YMD
		, nameL:fldval.PID_NameL
		, nameF:fldval.PID_NameF
		, name:fldval.PID_NameL strQ(fldval.PID_NameF,", ###")
		, mrn:fldval.PID_PatMRN
		, sex:(fldval.PID_sex~="F") ? "Female" : "Male"
		, DOB:parseDate(fldval.PID_DOB).MDY
		, monitor:monType
		, mon:monType
		, provider:prov
		, prov:prov
		, provname:provname
		, type:encType
		, loc:location
		, Account:fldval.ORC_ReqNum
		, accountnum:fldval.PID_AcctNum
		, encnum:fldval.PV1_VisitNum
		, order:fldval.ORC_ReqNum
		, accession:fldval.ORC_FillerNum
		, acct:location strQ(fldval.ORC_ReqNum,"_###") strQ(fldval.ORC_FillerNum,"-###")
		, UID:tobase(fldval.ORC_ReqNum RegExReplace(fldval.ORC_FillerNum,"[^0-9]"),36)
		, ind:indication
		, indication:indication
		, indicationCode:strQ(fldval.OBR_ReasonCode,"###") strQ(indCode,"###")
		, orderCtrl:fldval.ORC_OrderCtrl
		, ctrlID:fldval.MSH_CtrlID}
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
						
						tmpPrvFuzz := fuzzySearch(format("{:U}"							; degrees of fuzz for ptDem.Provider with clicked value
									, ptDem.Provider)
									, format("{:U}",tmpPrv))
						
						if (ptDem.Provider="") {
							ptDem.Provider := tmpPrv
							eventlog("MouseGrab provider empty --> " tmpPrv ".")
							
						} else if (tmpPrvFuzz>0.15) {									; names differ by more than 15%
							MsgBox, 4148
								, Provider already exists
								, % "Replace " ptDem.Provider "`n with `n" tmpPrv "?"
							IfMsgBox, Yes												; Check before replacing
							{
								eventlog("Replacing provider """ ptDem.Provider """ with """ tmpPrv """.")
								ptDem.Provider := tmpPrv
							}
						}																; otherwise ptDem.Provider exists and matches, so leave alone
						
					} else {															; no name clicked
						tmpPrv :=
						eventlog("MouseGrab provider empty.")
					}																	; tmpPrv will contain either clicked Provider or null
					
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
		mdTitle := RegExReplace(mdTitle,"\'","^")
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
	encDT := parseDate(ptDem.EncDate).YMD
	demBits := 0											; clear the error check
	fTxt := "`n	Verify that this is the valid patient information"
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
	;~ Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("Type","i)[a-z]+",1), Type
	;~ Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " cDefault", % ptDem["Type"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" fetchValid("Account","\d{4,}",1), Encounter #
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " vEncNum" " cDefault", % encNum
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" ((!(checkCrd(ptDem.Provider).fuzz=0)||!(ptDem.Provider))?"Red":"Default"), Ordering MD
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH  " cDefault", % ptDem["Provider"]
	Gui, fetch:Add, Button, % "x" fX1+10 " y" (fY += fYD) " h" fH+10 " w" fW1+fW2 " gfetchSubmit " ((demBits)?"Disabled":""), Submit!
	Gui, fetch:Show, AutoSize, Patient Demographics
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
	
	gotMD := false
	matchProv := checkCrd(ptDem.Provider)
	if (ptDem.Type~="i)(Inpatient|Observation|Emergency|Day Surg)") {					; encounter is an inpatient type (Inpt, ER, DaySurg, etc)
		encDT := ptDem.date
		ptDem.EncDate := niceDate(ptDem.date)											; set formatted EncDate
		gosub assignMD																	; find who recommended it from the Chipotle schedule
		ptDem.loc:="MAIN"
		eventlog(ptDem.Type " location. Provider assigned to " ptDem.Provider ".")
	}
	else if (matchProv.group="FELLOWS") {												; using fellow encounter
		ptDem.Fellow := matchProv.best
		eventlog("Fellow: " parseName(ptDem.fellow).firstlast)
		MsgBox, 262208, % parseName(ptDem.fellow).firstLast, Fellow-ordered monitor.`nMust also include the attending preceptor.
		gosub getMD
	}
	else if (matchProv.fuzz > 0.10) {													; Provider not recognized, ask!
		eventlog(ptDem.Provider " not recognized (" matchProv.fuzz ").")
		gosub getMD
		eventlog("Provider set to " ptDem.Provider ".")
	} 
	else if !(ptDem.Provider) {															; No provider? ask!
		gosub getMD
		eventlog("New provider field " ptDem.Provider ".")
	} 
	else {																				; Attending cardiologist recognized
		eventlog(ptDem.Provider " matches " matchProv.Best " (" (1-matchProv.fuzz)*100 ").")
		ptDem.Provider := matchProv.Best
	}
	
	while (gotMD=false)																	; do until we have while no confirmed cardiologist
	{
		MsgBox, 262180, Confirm attending Cardiologist, % ptDem.Provider
		IfMsgBox, Yes
		{
			break
		}
		gosub getMD
	}
	
	if (ptDem.acct="") {
		ptDem.acct := ptDem.loc strQ(fldval.ORC_ReqNum,"_###") strQ(fldval.ORC_FillerNum,"-###")
	}
	tmpCrd := checkCrd(ptDem.provider)													; Make sure we have most current provider
	ptDem.NPI := Docs[tmpCrd.Group ".npi",ObjHasValue(Docs[tmpCrd.Group],tmpCrd.best)]
	ptDem["Account"] := EncNum															; make sure array has submitted EncNum value
	FormatTime, EncDt, %EncDt%, MM/dd/yyyy												; and the properly formatted date 06/15/2016
	ptDem.EncDate := EncDt
	getDem := false																; done getting demographics
	Loop
	{
		if (ptDem.Indication) {													; loop until we have filled indChoices
			break
		}
		if (fetchquit=true) {
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

indGUIClose:
{
	Gui, ind:Destroy
	fetchQuit := true
	return
}

indSubmit:
{
	Gui, ind:Submit
	if InStr(indChoices,"OTHER",Yes) {
		InputBox(indOther, "Other", "Enter other indication","")
		indChoices := RegExReplace(indChoices,"OTHER", "OTHER - " indOther)
	}
	ptDem["Indication"] := RegExReplace(indChoices,"\|","; ")
	eventlog("Indications entered.")
	return
}

getDem:
{
	gosub fetchGUI																		; Grab it first
	WinWaitClose, Patient Demographics
	/*	Need to get demo validating code from fetchDem
	*/
	if (fetchQuit=true) {
		return
	}
	Loop
	{
		if (ptDem.Indication) {															; loop until we have filled indChoices
			break
		}
		if (fetchQuit=true) {
			break
		}
		gosub indGUI
		WinWaitClose, Enter indications
	}
	
	return
}

PrevGrab:
{
	Run, PrevGrab.exe
	return
}

enrollcheck(params) {
	global wq
	
	en := wq.selectSingleNode("//enroll" params)
	id := en.getAttribute("id")
	
; 	returns id if finds a match, else null
	return id																			
}

findWQid(DT:="",MRN:="",ser:="") {
/*	DT = 20170803
	MRN = 123456789
	ser = BodyGuardian Heart - BG12345, or Mortara H3+ - 12345
*/
	global wq
	
	if IsObject(x := wq.selectSingleNode("//enroll"
		. "[date='" DT "'][mrn='" MRN "']")) {												; Perfect match DT and MRN
	} else if IsObject(x := wq.selectSingleNode("//enroll"
		. "[dev='" ser "'][mrn='" MRN "']")) {												; or matches S/N and MRN
	} else if IsObject(x := wq.selectSingleNode("//enroll"
		. "[date='" DT "'][dev='" ser "']")) {												; or matches DT and S/N
	} else {
		x :=																				; anything else is null
	}

	return {id:x.getAttribute("id"),node:x.parentNode.nodeName}								; returns {id,node}; or null (error) if no match
}

cleanTempFiles() {
	thresh:=180
	
	Loop, files, tempfiles\*
	{
		filenm := A_LoopFileName
		FileGetTime, fileCDT, % "tempfiles\" filenm, C
		dtDiff := A_now
		dtDiff -= fileCDT, Days
		if (dtDiff<thresh) {															; skip younger files, default 180 days
			ct_skip ++
			continue
		}
		
		if RegExMatch(filenm,"\.csv$") {												; handle CSV files
			RegExMatch(filenm,"(\d{2}-\d{2}-\d{4})\.csv",v)
			dt := parseDate(v1)
			if (dt.date) {																; move if has a valid date
				dtStr := dt.yyyy dt.mm dt.dd
				DestDir := "tempfiles\archived\" dt.yyyy "\" dt.mm
				if !instr(FileExist(DestDir),"D") {						; 
					FileCreateDir, % DestDir
				}
				FileMove, % "tempfiles\" filenm, % DestDir "\" filenm
				ct_csv ++
				continue
			}
		} else {
			FileDelete, % "tempfiles\" filenm
			ct_other ++
		}
	}
	MsgBox, 262208, , % ""
		. "CSV files moved: " ct_csv "`n"
		. "Files deleted: " ct_other "`n"
		. "FIles skipped: " ct_skip
	return
}

checkMWUapp()
{
	global isDevt
	
	if (isDevt=true) {																	; In DEVT environment, skip loading MWU
		eventlog("isDevt=true, skip MWU load.")
		return
	}
	
	if !WinExist("ahk_exe WebUploadApplication.exe") {									; launch Mortara Upload app from site if not running
		wb := IEopen()
		sleep 500
		wb.Navigate("https://h3.preventice.com/WebUploadApplication.application")		; open direct link to WebUploadApplication.application
		wb := {}																		; disconnect the webbrowser object
		
		progress, y150,,Loading Mortara program...
		loop, 100																		; loop up to 30 seconds for window to appear
		{
			progress, % A_index
			if WinExist("Mortara Web Upload") {
				break
			}
			sleep 100
		}
		progress, off
	}
	
	return																	
}

MortaraUpload(tabnum="")
{
	global wq, mu_UI, ptDem, fetchQuit, MtCt, webUploadDir, user, isDevt, mwuPhase
	checkMachine()
	SetTimer, idleTimer, Off
	
	checkMWUapp()
	
	muWinID := WinExist("Mortara Web Upload")

	fetchQuit := false
	MtCt := ""
	mu_UI := MorUIgrab()
	muWinTxt := mu_UI.vis
	
	SerNum := substr(stregX(muWintxt,"Status.*?[\r\n]+",1,1,"Recorder S/N",1),-6)		; Get S/N on visible page
	SerNum := SerNum ? trim(SerNum," `r`n") : ""
	if (isDevt=true) {
		SerNum := "12345"
		;~ Tabnum := cMsgBox("DEVT MortaraUpload","Which tab?","Prepare|Transfer","Q")
	}
	if (SerNum="") {
		eventlog("No device attached, return to PhaseGUI.")
		return
	} else {
		eventlog("Device S/N " sernum " attached.")
	}
	
	if (Tabnum="Transfer") {															; TRANSFER RECORDING TAB
		eventlog("Transfer recording selected.")
		
		if (mwuPhase != Tabnum) {
			MsgBox, 262160, Mortara app selection, Switch the Mortara app tab to`n"Transfer Recording".`n`nClick "OK" to continue
			SetTimer, idleTimer, 500
			return
		}
		
		ptDem := Object()
		
		wuDir := {}
		Loop, files, % WebUploadDir "Data\*", D											; Get the most recently created Data\xxx folder
		{
			loopDate := A_LoopFileTimeModified
			loopName := A_LoopFileLongPath
			if (loopDate>=wuDir.Date) {
				wuDir.Date := loopDate
				wuDir.Full := loopName
			}
			wuDir.fullDir := loopDate ", " loopname "`n"
		}
		if (wuDir.Full="") {															; no transfer files found
			eventlog("No transfer files found.")
			MsgBox, 262160, Device error, No transfer files found!`n`nTry again.
			muPushButton(muWinID,"Back")
			return
		}
		wuDir.Short := strX(wuDir.Full,"\",0,1,"",0)									; transfer files found
		eventlog("Found WebUploadDir " wuDir.Short )
		FileReadLine, wuRecord, % wuDir.Full "\RECORD.LOG", 1
		FileReadLine, wuDevice, % wuDir.Full "\DEVICE.LOG", 1
		FileRead, wuConfig, % wuDir.Full "\CONFIG.SYS"
			wuConfig := substr(wuConfig,1,512)
			RegExMatch(wuConfig,"\D(\d{5,})\D+(\d{6,})?\D",t)
		if (t1) {
			wuDir.Ser := substr(t1,1-strlen(sernum))
			eventlog("wuDirSer " wuDir.Ser " from CONFIG.SYS")
		} else if (wuDevice) {
			wuDir.Ser := substr(wuDevice,-4)
			eventlog("wuDirSer " wuDir.Ser " from DEVICE.LOG")
		} else {
			eventlog("No S/N found.")
		}
		if (t2) {
			wuDir.MRN := t2
			eventlog("wuDirMRN " wuDir.MRN " from CONFIG.SYS")
		} else if (wuRecord) {
			wuDir.MRN := trim(RegExReplace(wuRecord,"i)Patient ID:"))
			eventlog("wuDirMRN " wuDir.MRN " from RECORD.LOG")
		} else {
			eventlog("No MRN found.")
		}
		if !(serNum=wuDir.Ser) {
			eventlog("Serial number mismatch.")
			eventlog(wuDir.fullDir)
			FileAppend, % A_now "|" A_UserName "|" A_ComputerName "|" serNum "`n", badSerNum.txt
			MsgBox, 262160, Device error, Device mismatch!`n`nTry again.
			muPushButton(muWinID,"Back")
			return
		}
		
		wq := new XML("worklist.xml")													; refresh WQ
		wqStr := "/root/pending/enroll[dev='Mortara H3+ - " SerNum "'][mrn='" wuDir.MRN "']"
		wqTR:=wq.selectSingleNode(wqStr)
		
		pt := readwq(wqTR.getAttribute("id"))
		ptDem["mrn"] := pt.mrn															; fill ptDem[] with values
		ptDem["loc"] := pt.site
		ptDem["date"] := pt.date
		ptDem["Account"] := RegExMatch(pt.acct,"([[:alpha:]]+)(\d{8,})",z) ? z2 : pt.acct
		ptDem["nameL"] := parseName(pt.name).last
		ptDem["nameF"] := parseName(pt.name).first
		ptDem["Sex"] := pt.sex
		ptDem["dob"] := pt.dob
		ptDem["Provider"] := pt.prov
		ptDem["Indication"] := pt.ind
		ptDem["loc"] := z1
		ptDem["wqid"] := wqTR.getAttribute("id")
		
		if IsObject(wqTR.selectSingleNode("acct")) {									; node exists, and valid
			eventlog("Found valid registration for " pt.name " " pt.mrn " " pt.date)
			MsgBox, 262193
				, Match!
				, % "Found valid registration for:`n   " pt.name "`n   " pt.mrn "`n   " pt.date "`n`nContinue?"
			IfMsgBox, Cancel
			{
				eventlog("Cancelled GUI.")
				muPushButton(muWinID,"Back")
				return
			}
		} 
		else if (wqTR.getAttribute("id")) {												; node exists, but not validated
			eventlog("Found unvalidated registration for " pt.name " " pt.mrn " " pt.date)
			MsgBox, 262193
				, Match?
				, % "Found unvalidated registration for:`n   " pt.name "`n   " pt.mrn "`n   " parseDate(pt.date).mdy "`n`nContinue?"
			IfMsgBox, Cancel
			{
				eventlog("Cancelled GUI.")
				muPushButton(muWinID,"Back")
				return
			}
		}
		else {																			; no matching node found
			eventlog("No registration found for " pt.name " " pt.mrn " " pt.date)
			MsgBox, 262193
				, No match
				, % "No registration found for:`n   " pt.name "`n   " pt.mrn "`n   " pt.date "`n`nContinue?"
			IfMsgBox, Cancel
			{
				eventlog("Cancelled GUI.")
				muPushButton(muWinID,"Back")
				return
			}
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
		eventlog(ptDem.MRN " " ptDem.nameL " study " ptDem.Date " uploaded to Preventice.")
		mwuPhase := ""
		MsgBox, 262208, Transfer, Successful data upload to Preventice.
	}
	
	if (Tabnum="Prepare") {																; PREPARE MEDIA TAB
		eventlog("Prepare media selected.")
		
		if (mwuPhase != Tabnum) {
			MsgBox, 262160, Mortara app selection, Switch the Mortara app tab to`n"Prepare Recorder Media".`n`nClick "OK" to continue
			SetTimer, idleTimer, 500
			return
		}
		
		if (ptDem.filename="") {
			MsgBox, 262160, Mortara app selection, Please reselect order from ORDERS tab.
			SetTimer, idleTimer, 500
			return
		}
		filein := ptDem.filename														; refresh ptDem and fldval from ORM
		processhl7(fileIn)																; because WQlist wipes out fldval
		ptDem:=parseORM()
		ptDem.filename := fileIn
		ptDem.Provider := ptDem.provname
		
		gosub getDem
		if (fetchQuit=true) {
			fetchQuit:=false
			eventlog("Cancelled getDem.")
			muPushButton(muWinID,"Back")
			return
		}
		getPatInfo()																	; grab remaining demographics for Preventice registration
		if (fetchQuit=true) {
			fetchQuit:=false
			eventlog("Cancelled getPatInfo.")
			muPushButton(muWinID,"Back")
			return
		}
		
		MorUIfill(mu_UI.PRct,muWinID)													; Fill UI fields from ptDem
		
		if (isDevt=false) {
			muPushButton(muWinID,"Set Clock...")										; Make sure clock button is set
			WinWaitClose, Set Recorder Time
			
			loop											
			{
				winget, x, ProcessName, A												; Dialog has no title
				if !instr(x,"WebUpload") {												; so find the WebUpload
					continue
				}
				WinGetText, x, A
				if (x="OK`r`n") {														; dialog that has only "OK`r`n" as the text
					WinGet, finOK, ID, A
					break
				}
			}
			Winwaitclose, ahk_id %finOK%												; Now we can wait until it is closed
		}
		
		InputBox(note, "Fedex", "`n`n`n`n Enter FedEx return sticker number","")
		if (note) {
			ptDem["fedex"] := note
			eventlog("Fedex number entered.")
		} else {
			eventlog("Fedex ignored.")
		}
		
		wq := new XML("worklist.xml")													; refresh WQ
		ptDem["muphase"] := "prepare"
		ptDem["hookup"] := "Office"
		muWqSave(SerNum)
		eventlog(ptDem["muphase"] ": " sernum " registered to " ptDem["mrn"] " " ptDem["nameL"] ".") 
		
		/*	This is just for Epic orders testing
		*/
		if (isDevt=true) {
			makeTestORU()
		}
		/*
		*/
		
		removeNode("/root/orders/enroll[@id='" ptDem.uid "']")
		writeOut("root","orders")
		wq := new XML("worklist.xml")
		FileMove, % ptDem.filename, .\tempfiles, 1
		
		makePreventiceORM()
		mwuPhase := ""
	}
	
	return
}

muPushButton(muWinID,btn) {
	WinActivate, ahk_id %muWinID%
	sleep 500
	ControlGet, clkbut, HWND,, %btn%
	sleep 200
	ControlClick,, ahk_id %clkbut%,,,,NA
	
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
			
			WriteSave(wq)
		eventlog("Device " sernum " reg to " enName " - " enMRN " on " enDate ", moved to DONE list.")
	}
	wq := new XML("worklist.xml")
	
	id := ptDem.UID
	ptDem["model"] := "Mortara H3+"
	ptDem["ser"] := sernum
	ptDem["dev"] := ptDem.model " - " sernum
	ptDem["wqid"] := id
	ptDem["date"] := parsedate(ptDem["EncDate"]).YMD									; make sure ptDem.date in proper format
	
	wq.addElement("enroll","/root/pending",{id:id})
	ptDem.newID := "/root/pending/enroll[@id='" id "']"
	wq.addElement("date",ptDem.newID,ptDem.date)
	wq.addElement("name",ptDem.newID,ptDem.name)
	wq.addElement("mrn",ptDem.newID,ptDem.mrn)
	wq.addElement("sex",ptDem.newID,ptDem.sex)
	wq.addElement("dob",ptDem.newID,ptDem.dob)
	wq.addElement("dev",ptDem.newID,ptDem.dev)
	if (ptDem.fellow) {
		wq.addElement("fellow",ptDem.newID,ptDem.fellow)
	}
	wq.addElement("prov",ptDem.newID,ptDem.Provider)
	wq.addElement("site",ptDem.newID,ptDem.loc)										; need to transform site abbrevs
	wq.addElement("acct",ptDem.newID,ptDem.acct)
	wq.addElement("order",ptDem.newID,ptDem.order)
	wq.addElement("accession",ptDem.newID,ptDem.accession)
	wq.addElement("accountnum",ptDem.newID,ptDem.accountnum)
	wq.addElement("encnum",ptDem.newID,ptDem.encnum)
	wq.addElement("ind",ptDem.newID,ptDem.indication)
	if (ptDem.fedex) {
		wq.addElement("fedex",ptDem.newID,ptDem.fedex)
	}
	wq.addElement(ptDem["muphase"],ptDem.newID,{user:A_UserName},A_now)
	
	filedelete, .lock
	writeOut("/root/pending","enroll[@id='" id "']")
	wq := new XML("worklist.xml")
	
	return
}

MorUIgrab() {
	local visTxt, WinText, Wintab
		, mx, my, mw, mh
	
	id := WinExist("Mortara Web Upload")
	DetectHiddenText, off
	WinGetText, visTxt, ahk_id %id%											; Should only check visible window
	DetectHiddenText, on
	q := Object()
	WinGet, WinText, ControlList, ahk_id %id%
	ControlGet , Wintab, Tab,
		, WindowsForms10.SysTabControl32.app.0.33c0d9d1
		, ahk_id %id%

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
	q.tab := Wintab
	q.vis := vistxt
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

makePreventiceORM() {
	global wq, ptDem, fetchQuit, hl7out, path, indCodes, sitesCode, sitesFacility
	
	hl7time := A_Now
	hl7out := Object()
	buildHL7("MSH"
		,{1:"^~\&"
		, 2:"TRRIQ"
		, 3:sitesCode
		, 4:sitesFacility
		, 5:"PREVENTICE"
		, 6:hl7time
		, 7:"TECH"
		, 8:"ORM^O01"
		, 9:ptDem["wqid"]
		, 10:"T"
		, 11:"2.3" })
	
	buildHL7("PID"
		,{2:ptDem.MRN
		, 3:ptDem.MRN
		, 5:ptDem.nameL "^" ptDem.nameF . strQ(ptDem.nameMI,"^###")
		, 7:parseDate(ptDem.dob).YMD
		, 8:substr(ptDem.sex,1,1)
		, 11:ptDem.Addr1 "^" ptDem.Addr2 "^" ptDem.city "^" ptDem.state "^" ptDem.zip
		, 13:ptDem.phone
		, 18:ptDem.account })
	
	tmpPrv := parseName(ptDem.provider)
	buildHL7("PV1"
		,{2:ptDem.type
		, 3:ptDem.loc
		, 7:ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, 8:ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, 19:ptDem.wqid })
	
	buildHL7("IN1"
		,{2:"N/A"
		, 4:"Seattle Childrens - GB" ;"Insurance Company Name"
		, 16:ptDem.parentL "^" ptDem.parentF
		, 17:"Legal Guardian"
		, 18:parseDate(ptDem.dob).YMD })
	
	buildHL7("ORC",{2:""})
	
	buildHL7("OBR"
		,{2:ptDem.wqid
		, 4:strQ((ptDem.model~="Mortara") ? 1 : "","Holter^Holter")
			. strQ((ptDem.model~="Heart") ? 1 : "","CEM^CEM")
			. strQ((ptDem.model~="Mini") ? 1 : "","Holter^Holter")
		, 7:hl7time
		, 11:"ANCILLARY"
		, 16:ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, 17:"206-987-2015" })
	
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
			,{3:indIdx
			, 4:indSeg })
	}
	
	buildHL7("OBX"
		,{2:"ST"
		, 3:"12915^Service Type"
		, 5:strQ((ptDem.model~="Mortara") ? 1 : "","Holter")
			. strQ((ptDem.model~="Heart") ? 1 : "","CEM")
			. strQ((ptDem.model~="Mini") ? 1 : "","Holter") })
	
	buildHL7("OBX"
		,{2:"ST"
		, 3:"12916^Device"
		, 5:ptDem.model })
	
	buildHL7("OBX"
		,{2:"ST"
		, 3:"12919^Serial Number"
		, 5:ptDem.ser })
	
	buildHL7("OBX"
		,{2:"ST"
		, 3:"12917^Hookup Location"
		, 5:ptDem.Hookup })
	
	buildHL7("OBX"
		,{2:"ST"
		, 3:"12918^Deploy Duration (In Days)"
		, 5:(ptDem.model~="Mortara" ? "1" : "")
			. (ptDem.model~="Heart" ? "30" : "")
			. (ptDem.model~="Mini" ? strQ(ptDem.HolterDuration,"###","14") : "") })
	
	fileNm := ptDem.nameL "_" ptDem.nameF "_" ptDem.mrn "-" hl7time ".txt"
	FileAppend, % hl7Out.msg, % ".\tempfiles\" fileNm
	FileCopy, % ".\tempfiles\" fileNm , % path.PrevHL7out . fileNm
	eventlog("Preventice registration completed: " fileNm)
	MsgBox, 262208, Preventice registration, Successful device registration!
	return
}

BGregister(type) {
	global wq, ptDem, fetchQuit, isDevt
	SetTimer, idleTimer, Off
	
	typeLong := (type="BGH" ? "BodyGuardian Heart" : "") . (type="BGM" ? "BodyGuardian Mini EL" : "")
	
	tmp:=CMsgBox(ptDem.Monitor
		, "Register type`n`n" typeLong
			. (type="BGH" ? "`n30-day Event Recorder" : "")
			. (type="BGM" ? "`nExtended Holter (3-14 day)" : "")
		, "Yes|No"
		, "Q", "V", 
		, (type="BGH" ? ".\files\BGHeart.png" : "") 
		. (type="BGM" ? ".\files\BGMini.png" : "") )
	if (tmp!="Yes") {
		return
	}
	
	if (type="BGM") {
		tmp:=CMsgBox("Extended Holter duration"
			, "Select expected duration of recording"
			, "3 days|7 days|14 days"
			, "Q")
		if (tmp="xClose") {
			return
		}
		ptDem.HolterDuration := strX(tmp,"",1,0," ",1,1)
	}
	
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
	if (i="xClose") {
		eventlog("Cancelled delivery type.")
		return
	}
	if (i="Home") {
		ptDem["hookup"] := "Home"
		ptDem["model"] := typeLong
		eventlog(type " home registration for " ptDem["mrn"] " " ptDem["nameL"] ".") 
	} 
	else {																				; either Office or [X]
		ptDem["hookup"] := "Office"
		ptDem.ser := selectDev()														; need to grab a ser num from inventory
		if (ptDem.ser="") {
			eventlog("Cancelled selectDev.")
			return
		}
		ptDem.model := wq.selectSingleNode("/root/inventory/dev[@ser='" ptDem.ser "']").getAttribute("model")
		
		if !(ptDem.model) {																; Types in an ad hoc number
			i := cMsgBox("Recorder type","Which recorder?","BodyGuardian Heart|BodyGuardian Mini EL")
			if (i="xClose") {
				eventlog("Cancelled ad hoc S/N.")
				return
			} else {
				ptDem.model := i
				eventlog("User typed ad hoc S/N " ptDem.ser ", type " i ".")
			}
		}
		
		if (ptDem.model != typeLong) {													; Selects mismatched device
			MsgBox, 262161, SELECTION MISMATCH
				, % "Selected device " ptDem.model
				. "`ndoes not match `n"
				. "ordered device " typeLong "."
				. "`n`nAre you sure you want to proceed?"
			IfMsgBox, OK
			{
				selReason := cMsgBox("Order override","Reason to override order:"
					, "Provider requests change|"
					. "Inventory problem|"
					. "Wrong device"
					, "Q")
				if (selReason="xClose") {
					MsgBox Cancelled order processing
					eventlog("Cancelled user override.")
					return
				}
				eventlog("User override '" selReason "'. " typeLong " => " ptDem.model ".")
			}
			else {
				eventlog("Selection mismatch cancelled.")
				return
			}
		}
		
		removeNode("/root/inventory/dev[@ser='" ptDem.ser "']")							; take out of inventory
		writeOut("/root","inventory")
		eventlog(ptDem.ser " registered to " ptDem["mrn"] " " ptDem["nameL"] ".") 
	}
	wq := new XML("worklist.xml")														; refresh WQ
	bgWqSave(ptDem.ser)																; write to worklist.xml
	eventlog(type " " ptDem.ser " registered to " ptDem.mrn " " ptDem.nameL ".")
	
	removeNode("/root/orders/enroll[@id='" ptDem.uid "']")
	writeOut("root","orders")
	FileMove, % ptDem.filename, .\tempfiles, 1
		
	makePreventiceORM()
	
	/*	This is just for Epic orders testing
	*/
	if (isDevt=true) {
		makeTestORU()
	}
	/*
	*/
	
	return
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
		if !(choice~="^(BG)?\d{7}$") {													; ignore if doesn't match full ser num
			return
		}
		Gui, dev:Destroy
		return
	}

}

getPatInfo() {
/*	Parse guardians from Epic NK1 segments
*/
	global wq, ptDem, fetchQuit, fldval
	
	ptDem.phone := formatPhone(fldval.PID_phone)														; get phone num from PID
	
;	Now separate the "Family contact" members, grab relevant contact info from each parsed line
	relStr := "FAT|MOT|FOS|GPR|AOU|STP|INS"
	rel := Object()
	
	loop
	{
		i := A_Index
		pre := "NK1_" i "_"
		name := fldval[pre "NameL"] . strQ(fldval[pre "NameF"],", ###")
		if (name="") {
			break
		}
		rel[i] := object()
		rel[i].name := name
		rel[i].relation := fldval[pre "Relation"]
		tmp := segField(fldval[pre "Phone"],"num^type^equipment")
		rel[i].phoneHome := formatPhone(tmp.selectSingleNode("//idx[equipment/text()='HOME']/num").text)
		rel[i].phoneMobile := formatPhone(tmp.selectSingleNode("//idx[equipment/text()='MOBILE']/num").text)
		tmp := fldval[pre "Role"]
		rel[i].lives := instr(tmp,"Y^LW") ? true : false
		rel[i].legal := instr(tmp,"Y^LG") ? true : false
		rel[i].addr := strQ(fldval[pre "Addr1"],"###`n")
			. strQ(fldval[pre "Addr2"],"###`n")
			. strQ(strQ(fldval[pre "City"],"###") strQ(fldval[pre "State"],", ###") strQ(fldval[pre "Zip"]," ###"),"###`n")
	}
		
;	Filter out contacts who are not likely guarantors or parents
	loop, % rel.MaxIndex()
	{
		i := A_index
		if (rel[i].lives = true) {
			ptDem.livesaddr := rel[i].addr
			continue																	; keep if "Lives here" is true
		}
		if (rel[i].guardian = true) {
			continue																	; keep if is guardian
		}
		if ((rel[i].addr="") && (rel[i].phone="")) {
			rel.Delete(i)																; remove entries with no address or phone
			continue
		}
		rel.Delete(i)																	; remove anyone who doesn't match
	}
	
;	Generate parent name menu for cmsgbox selection
	if (rel.MaxIndex() > 1) {
		loop, % rel.MaxIndex()
		{
			nm .= A_index ") " rel[A_index].name "|"
		}
		eventlog("Multiple potential parent matches (" rel.MaxIndex() ").")
		q := cmsgbox("Parent","Who is the guarantor?",trim(nm,"|"))
		if (q="xClose") {
			eventlog("Quit registration at parent selection.")
			fetchQuit:=true
			return
		}
		choice := strX(q,"",1,1,")",1,1)
	} else {
		choice := 1
	}
	eventlog("Parent selection " choice ": " rel[choice].Name "|" rel[choice].livesaddr)
	
	ptDem.parent := rel[choice].Name
	ptDem.parentL := parseName(ptDem.parent).last
	ptDem.parentF := parseName(ptDem.parent).first
	
	ptDem.phone := (rel[choice].phoneHome) 
		? rel[choice].phoneHome
		: (rel[choice].phoneMobile)
			? rel[choice].phoneMobile
			: ""
	
	if (rel[choice].addr="") {
		rel[choice].addr := ptDem.livesaddr
	}
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
	if (ptDem.addr1="") {
		InputBox(addr1, "Registration requires street address","`n`nEnter valid street address","")
		InputBox(addr2, "Registration requires street address","`n`nEnter city", ptDem.city)
		if (addr1) {
			ptDem.addr1 := addr1
			eventlog("Entered street address.")
		} else {
			fetchQuit := true
			return
		}
	}
	if (ptDem.addr1~="i)^P[\. ]+?O[\. ]+?Box") {
		InputBox(addr1, "Cannot use P.O. Box","`n`nEnter valid street address","")
		InputBox(addr2, "Cannot use P.O. Box","`n`nEnter city", ptDem.city)
		if (addr1) {
			ptDem.addr1 := addr1
			eventlog("Replaced PO box with street address.")
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
		. strQ(ptDem.fellow,"Fellow: ###`n")
		. "Provider: " ptDem.provider "`n"
		. "Encounter date: " ptDem.encDate "`n"
		. "Site: " ptDem.loc
	IfMsgBox, Yes
	{
		eventlog("Selected parent " ptDem.parentL ", " ptDem.parentF)
		eventlog("Accepted patient address info. " ptDem.addr1 " | " strQ(ptDem.addr2,"### | ") ptDem.city " | " ptDem.state " " ptDem.zip)
		fetchQuit := false
	} else {
		fetchQuit := true
	}
	return
}

bgWqSave(sernum) {
	global wq, ptDem, user, sitesLong
	
	id := ptDem.UID
	ptDem["dev"] := ptDem.model " - " sernum
	ptDem["wqid"] := id
	ptDem["date"] := parsedate(ptDem["EncDate"]).YMD									; make sure ptDem.date in proper format
	
	wq.addElement("enroll","/root/pending",{id:id})
	ptDem.newID := "/root/pending/enroll[@id='" id "']"
	wq.addElement("date",ptDem.newID,ptDem.date)
	wq.addElement("name",ptDem.newID,ptDem.name)
	wq.addElement("mrn",ptDem.newID,ptDem.mrn)
	wq.addElement("sex",ptDem.newID,ptDem.sex)
	wq.addElement("dob",ptDem.newID,ptDem.dob)
	wq.addElement("dev",ptDem.newID,ptDem.dev)
	if (ptDem.fellow) {
		wq.addElement("fellow",ptDem.newID,ptDem.fellow)
	}
	wq.addElement("prov",ptDem.newID,ptDem.Provider)
	wq.addElement("site",ptDem.newID,ptDem.loc)										; need to transform site abbrevs
	wq.addElement("acct",ptDem.newID,ptDem.acct)
	wq.addElement("order",ptDem.newID,ptDem.order)
	wq.addElement("accession",ptDem.newID,ptDem.accession)
	wq.addElement("accountnum",ptDem.newID,ptDem.accountnum)
	wq.addElement("encnum",ptDem.newID,ptDem.encnum)
	wq.addElement("ind",ptDem.newID,ptDem.indication)
	if (ptDem.fedex) {
		wq.addElement("fedex",ptDem.newID,ptDem.fedex)
	}
	wq.addElement("register",ptDem.newID,{user:A_UserName},A_now)
	
	writeOut("/root/pending","enroll[@id='" id "']")
	
	return
}

moveHL7dem() {
/*	Populate fldVal["dem-"] with data from hl7 first, and wqlist (if missing)
*/
	global fldVal, obxVal
	
	name := parseName(fldval.name)
	fldVal["dem-Name_L"] := strQ(obxVal["PID_NameL"],"###",RegExReplace(name.last,"\^","'"))		; replace [^] with [']
	fldVal["dem-Name_F"] := strQ(obxVal["PID_NameF"],"###",RegExReplace(name.first,"\^","'"))
	fldVal["dem-Name"] := fldVal["dem-Name_L"] strQ(fldVal["dem-Name_F"],", ###")
	fldVal["dem-MRN"] := strQ(obxVal["PID_PatMRN"],"###",fldval.MRN)
	fldVal["dem-DOB"] := strQ(obxVal["PID_DOB"],niceDate(obxVal["PID_DOB"]),fldval.DOB)
	fldVal["dem-Sex"] := strQ(obxVal["PID_Sex"],(obxVal["PID_Sex"]~="F") ? "Female" : "Male",fldval.Sex)
	fldVal["dem-Indication"] := strQ(obxVal.Indications,"###",fldval.ind)
	fldVal["dem-Site"] := fldVal.site
	fldVal["dem-Billing"] := strQ(RegExReplace(fldVal.Accession,"[[:alpha:]]"),"###",RegExReplace(fldVal.acct,"[[:alpha:]]"))
	fldVal["dem-Ordering"] := strQ(fldval.fellow,"###",fldval.prov)
	fldVal["dem-Ordering"] := strQ(fldval["dem-Ordering"],"###",filterProv(obxVal["PV1_AttgNameF"] " " obxVal["PV1_AttgNameL"]).name)
	fldval["dem-Device_SN"] := strX(fldval.dev," ",0,1,"",0,0)

	return
}

ProcessHl7PDF:
{
/*	Associate fldVal data with extra metadata from extracted PDF, complete final CSV report, handle files
*/
	fileNam := RegExReplace(fldVal.Filename,"i)\.pdf")									; fileNam is name only without extension, no path
	fileIn := path.PrevHL7in fldVal.Filename											; fileIn has complete path \\childrens\files\HCCardiologyFiles\EP\HoltER Database\Holter PDFs\steve.pdf
	
	if (fileNam="") {																	; No PDF extracted
		eventlog("No PDF extracted.")
		progress, off
		MsgBox No PDF extracted!
		return
	}
	
	RunWait, .\files\pdftotext.exe -l 2 "%fileIn%" "%filenam%.txt",,min					; convert PDF pages 1-2 with no tabular structure
	FileRead, newtxt, %filenam%.txt														; load into newtxt
	FileDelete, %filenam%.txt
	StringReplace, newtxt, newtxt, `r`n`r`n, `r`n, All									; remove double CRLF
	FileAppend % newtxt, %filenam%.txt													; create new tempfile with result, minus PDF
	FileMove %filenam%.txt, .\tempfiles\*, 1											; move a copy into tempfiles for troubleshooting
	FileAppend % fldval.hl7, %filenam%_hl7.txt											; create a copy of hl7 file
	FileMove %filenam%_hl7.txt, .\tempfiles\*, 1										; move into tempfiles for troubleshooting
	
	progress, off
	type := fldval["OBR_TestCode"]														; study report type in OBR_testcode field
	if (type~="CEM|EOS") {
		gosub Event_BGH_Hl7
	} else if (ftype="MINI") {
		gosub Holter_BGM_HL7
	} else if (type="Holter") {
		gosub Holter_Pr_Hl7
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
	RunWait, .\files\pdftotext.exe -l 2 -table -fixed 3 "%fileIn%" "%filenam%.txt",,min			; convert PDF pages 1-2 to txt file
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
	} else if (instr(newtxt,"Global Instrumentation LLC")) {							; BG Mini extended Holter
		gosub Holter_BGM
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
	filenameOut := RegExReplace(filenameOut,"\^","'")										; convert [^] back to [']
	
	/*	Save hl7Out result
	*/
	tmpFile := ".\tempfiles\"																; HL7 for tempfiles,
		. "TRRIQ_ORU_" 																		; to copy to RawHL7 (for Access use)
		. fldval["dem-Name_L"] "_" 
		. tmpDate.YMD "_"
		. "@" fldval["wqid"] ".hl7"
	FileDelete, % tmpFile
	FileAppend, % hl7Out.msg, % tmpFile														; copy ORU hl7 to tempfiles
	FileCopy, % tmpFile, % path.EpicHL7out													; create copy in RawHL7
	
	/*	Save CSV in tempfiles, and copy to Import folder
	*/
	FileDelete, .\tempfiles\%fileNameOut%.csv												; clear any previous CSV
	FileAppend, %fileOut%, .\tempfiles\%fileNameOut%.csv									; create a new CSV in tempfiles
	
	impSub := (monType~="BGH") ? "EventCSV\" : "HolterCSV\"										; Import subfolder Event or Holter
	FileCopy, .\tempfiles\%fileNameOut%.csv, % path.import impSub "*.*", 1					; copy CSV from tempfiles to importFld\impSub
	
	/*	Copy PDF to OnBase
	*/
	onbaseFile := path.OnBase																; PDF for OnBase
		. "TRRIQ_" 
		. fldval["order"] "_" 
		. tmpDate.YMD "_" 
		. fldval["dem-Name_L"] "_" 
		. fldval["dem-MRN"] ".pdf"
	
	fileHIM := FileExist(fileIn "-sh.pdf")													; filename for OnbaseDir
			? fileIn "-sh.pdf"																; prefer shortened if it exists
			: fileIn
	
	FileCopy, % fileHIM, % onbaseFile, 1													; Copy to OnbaseDir
	
	/*	Copy PDF to HolterPDF folder and archive
	*/
	FileCopy, % fileIn, % path.holterPDF "Archive\" filenameOut ".pdf", 1					; Copy the original PDF to holterDir Archive
	FileCopy, % fileHIM, % path.holterPDF filenameOut "-short.pdf", 1						; Copy the shortened PDF, if it exists
	FileDelete, %fileIn%																	; Need to use Copy+Delete because if file opened
	FileDelete, %fileIn%-sh.pdf																;	was never completing filemove
	;~ FileDelete, % path.PrevHL7in fileNam ".hl7"											; We can delete the original HL7, if exists
	FileMove, % path.PrevHL7in fileNam ".hl7", .\tempfiles\%fileNam%.hl7
	eventlog("Move files '" fileIn "' -> '" filenameOut)
	
	/*	Append info to fileWQ (probably obsolete in Epic)
	*/
	fileWQ := ma_date "," user "," 															; date processed and MA user
			. """" fldval["dem-Ordering"] """" ","											; extracted provider
			. """" fldval["dem-Name_L"] ", " fldval["dem-Name_F"] """" ","					; CIS name
			. """" fldval["dem-MRN"] """" ","												; CIS MRN
			. """" fldval["dem-Test_date"] """" ","											; extracted Test date (or CIS encounter date if none)
			. """" fldval["dem-Test_end"] """" ","											; extracted Test end
			. """" fldval["dem-Site"] """" ","												; CIS location
			. """" fldval["dem-Indication"] """" ","										; Indication
			. """" monType """" ; ","														; Monitor type
			. "`n"
	FileAppend, %fileWQ%, .\logs\fileWQ.csv													; Add to logs\fileWQ list
	FileCopy, .\logs\fileWQ.csv, % path.chip "fileWQ-copy.csv", 1
	
	setwqupdate()
	wq := new XML("worklist.xml")
	moveWQ(fldval["wqid"])																	; Move enroll[@id] from Pending to Done list
	
	if (RegExMatch(fldval["dem-Ordering"], "Oi)(Chun|Salerno|Seslar)"))  {
		enc_MD := parseName(fldval["dem-Ordering"]).init
		tmp := httpComm("read&to=" enc_MD)
		eventlog("Notification email " tmp " to " enc_MD)
	}

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
	} else {																			; no record exists (enrollment never captured, or Zio)
		id := makeUID()																	; create an id
		wq.addElement("enroll","/root/done",{id:id})									; in </root/done>
		newID := "/root/done/enroll[@id='" id "']"
		wq.addElement("date",newID,parseDate(fldval["dem-Test_date"]).YMD)				; add these to the new done node
		wq.addElement("name",newID,fldval["dem-Name"])
		wq.addElement("mrn",newID,fldval["dem-MRN"])
		wq.addElement("done",newID,{user:A_UserName},A_Now)
		eventlog("No wqid. Saved new DONE record " fldval["dem-MRN"] ".")
	}
	writeSave(wq)
	
	FileDelete, .lock
	
	return
}

wqSetVal(id,node,val) {
	global wq
	
	newID := "/root/pending/enroll[@id='" id "']"
	k := wq.selectSingleNode(newID "/" node)
	if (k.text) and (val="") {															; don't overwrite an existing value with null
		return
	}
	val := RegExReplace(val,"\'","^")													; make sure no val ever contains [']
	
	if IsObject(k) {
		wq.setText(newID "/" node,val)
	} else {
		wq.addElement(node,newID,val)
	}
	
	return
}


getMD:
{
	gotMD := false
	Gui, fetch:Hide
	InputBox(ed_Crd, "Assign attending cardiologist","","")								; no call schedule for that day, must choose
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
		} else {
			gosub getMD											; don't agree? try again
		}
	}
	gotMD := true
	eventlog("Cardiologist " ptDem.Provider " entered.")
	return
}	

assignMD:
{
	if !(ptDem.date) {																	; must have a date to figure it out
		return
	}
	
	y := new XML(".\files\call.xml")													; get most recent schedule
	yNode := "//call[@date='" ptDem.date "']"
	ymatch := (ptDem.loc~="ICU") 
		? y.selectSingleNode(yNode "/ICU_A").text										; if order came from ICU
		: y.selectSingleNode(yNode "/Ward_A").text										; everything else
	if !(ymatch) {
		ymatch := y.selectSingleNode(yNode "/PM_We_A").text								; no match, must be a weekend
	}
	if (ymatch) {
		inptMD := checkCrd(ymatch)
		if (inptMD.fuzz<0.15) {															; close enough match
			ptDem.Provider := inptMD.best
			eventlog("Cardiologist autoselected " ptDem.Provider )
			return
		}
	} else {
		gosub getMD																		; when all else fails, ask
	}
return
}

epRead() {
	global y, path, user, ma_date, fldval
	
	y := new XML(".\files\call.xml")
	dlDate := A_Now
	FormatTime, dlDay, %dlDate%, dddd
	if (dlDay="Friday") {
		dlDate += 3, Days
	}
	FormatTime, dlDate, %dlDate%, yyyyMMdd
	
	RegExMatch(y.selectSingleNode("//call[@date='" dlDate "']/EP").text, "Oi)(Chun|Salerno|Seslar)", ymatch)
	if !(ep := ymatch.value()) {
		ep := cmsgbox("Electronic Forecast not complete","Which EP on Monday?","Chun|Salerno|Seslar","Q")
		if (ep="xClose") {
			eventlog("Elec Forecast not complete. Quit EP selection.")
		}
		eventlog("Reading EP assigned to " ep ".")
	}
	
	if (RegExMatch(fldval["dem-Ordering"], "Oi)(Chun|Salerno|Seslar)", epOrder))  {
		ep := epOrder.value()
	}
	fldval["dem-Reading"] := ep
	
	FormatTime, ma_date, A_Now, MM/dd/yyyy
	fieldcoladd("","EP_read",ep)
	fieldcoladd("","EP_date",niceDate(dlDate))
	fieldcoladd("","MA",user)
	fieldcoladd("","MA_date",ma_date)
	fieldcoladd("TRRIQ","UID",fldval.wqid)
	fieldcoladd("TRRIQ","order",fldval.order)
	fieldcoladd("TRRIQ","accession",fldval.accession)
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
	
	if (fldval["hrd-Total_beats"]="") {													; apparently no DDE present
		rateStat := stregX(newtxt,"(\R)ALL BEATS",1,0,"(\R)PAUSES",1) "<<<"
		if RegExMatch(rateStat, "Minimum HR.*?Average HR") {
			fields[1] := ["Total QRS", "Normal Beats"
				, "Minimum HR","Maximum HR","Average HR","Tachycardia"
				, "Longest Tachycardia","Fastest Tachycardia","Longest Bradycardia","Slowest Bradycardia","<<<"]
			labels[1] := ["Total_beats","null"
				, "Min","Max","Avg","null"
				, "Longest_tachy","Fastest","Longest_brady","Slowest","null"]
		} else {
			fields[1] := ["Total QRS", "Normal Beats"
				, "Minimum HR","Maximum HR","\R","Average HR","\R"
				, "Longest Tachycardia","Fastest Tachycardia","Longest Bradycardia","Slowest Bradycardia","<<<"]
			labels[1] := ["Total_beats","null"
				, "Min","Max","null","Avg","null"
				, "Longest_tachy","Fastest","Longest_brady","Slowest","null"]
		}
		fieldvals(rateStat,1,"hrd")
		
		rateStat := stregX(newtxt,"(\R)VENTRICULAR ECTOPY",1,0,"PACED|SUPRAVENTRICULAR ECTOPY",1)
		fields[2] := ["Ventricular Beats","Singlets","Couplets","Runs","Fastest Run","Slowest Run","Longest Run","R on T Beats"]
		labels[2] := ["Total","SingleVE","Couplets","Runs","Fastest","Slowest","Longest","R on T"]
		fieldvals(rateStat,2,"ve")
		
		rateStat := stregX(newtxt,"(\R)SUPRAVENTRICULAR ECTOPY",1,0,"(\R)OTHER RHYTHM EPISODES|(\R)RR VARIABILITY",1) "<<<"
		fields[3] := ["Supraventricular Beats","Aberrant Beats","Singlets","Pairs","Runs","Fastest Run","Slowest Run","Longest Run","SVE"
			, "Pauses .* ms","Longest RR","<<<"]
		labels[3] := ["Total","Aberrant","Single","Pairs","Runs","Fastest","Slowest","Longest","null"
			, "Pauses","LongestRR","null"]
		fieldvals(rateStat,3,"sve")
		
		eventlog("<<< Missing DDE, parsed from extracted PDF >>>")
	}
	
	if !(fldval.acct) {																	; fldval.acct exists if Holter has been processed
		gosub checkProc																	; get valid demographics
		if (fetchQuit=true) {
			return
		}
	}
	
	fieldsToCSV()
	
	FileGetSize, fileInSize, % path.PrevHL7in fldval.Filename
	if (fileInSize > 2000000) {															; probably a full disclosure PDF
		shortenPDF(fullDisc)															; generates .pdf and sh.pdf versions
	} 
	else loop {																			; just a short PDF
		if (findFullPDF(fldval.wqid)) {
			eventlog("Full disclosure PDF found.")
			break																		; found matching full disclosure, exit loop
		}
		else if (fldval.MSH_CtrlID~="EPIC") {
			eventlog("Epic test patient. No full disclosure PDF.")
			break
		}
		else {
			eventlog("Full disclosure PDF not found.")
			
			msg := cmsgbox("Missing full disclosure PDF"
				, fldval["dem-Name_L"] ", " fldval["dem-Name_F"] "`n`n"
				. "Download from ftp.eCardio.com site`n"
				. "then click [Retry].`n`n"
				. "If full disclosure PDF not available,`n"
				. "click [Email] to send a message to Preventice."
				, "Retry|Email|Cancel"
				, "E", "V")
			if (msg="Retry") {
				findFullPDF()
				continue
			}
			if (msg~="Cancel|Close|xClose") {
				FileDelete, % fileIn
				eventlog("Refused to get full disclosure. Extracted PDF deleted.")
				Exit																	; either Cancel or X, go back to main GUI
			}
			if (msg="Email") {
				progress,100 ,,Generating email...
				Eml := ComObjCreate("Outlook.Application").CreateItem(0)				; Create item [0]
				Eml.BodyFormat := 2														; HTML format
				
				Eml.To := "HolterNotificationGroup@preventice.com"
				Eml.cc := "EkgMaInbox@seattlechildrens.org; terrence.chun@seattlechildrens.org"
				Eml.Subject := "Missing full disclosure PDF"
				Eml.Display																; Display first to get default signature
				Eml.HTMLBody := "Please upload the full disclosure PDF for " fldval["dem-Name_L"] ", " fldval["dem-Name_F"] 
					. " MRN#" fldval["dem-MRN"] " study date " fldval["dem-Test_date"]
					. " to the eCardio FTP site.<br><br>Thank you!<br>"
					. Eml.HTMLBody														; Prepend to existing default message
				progress, off
				ObjRelease(Eml)															; or Eml:=""
				eventlog("Email sent to Preventice.")
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
	
	if (monType~="PR|Zio|Mini") {
		tabs := "dem-Name_L	dem-Name_F	dem-Name_M	dem-MRN	dem-DOB	dem-Sex(NA)	dem-Site	dem-Billing	dem-Device_SN	dem-VOID1	"
			. "dem-Hookup_tech	dem-VOID2	dem-Meds	dem-Ordering	dem-Scanned_by	dem-Reading	"
			. "dem-Test_date	dem-Scan_date	dem-Hookup_time	dem-Recording_time	dem-Analysis_time	dem-Indication	dem-VOID3	"
			. "hrd-Total_beats(0)	hrd-Min(0)	hrd-Min_time	hrd-Avg(0)	hrd-Max(0)	hrd-Max_time	hrd-HRV	"
			. "ve-Total(0)	ve-Total_per(0)	ve-Runs(0)	ve-Beats(0)	ve-Longest(0)	ve-Longest_time	ve-Fastest(0)	ve-Fastest_time	"
			. "ve-Triplets(0)	ve-Couplets(0)	ve-SinglePVC(0)	ve-InterpPVC(0)	ve-R_on_T(0)	ve-SingleVE(0)	ve-LateVE(0)	"
			. "ve-Bigem(0)	ve-Trigem(0)	ve-SVE(0)	sve-Total(0)	sve-Total_per(0)	sve-Runs(0)	sve-Beats(0)	"
			. "sve-Longest(0)	sve-Longest_time	sve-Fastest(0)	sve-Fastest_time	sve-Pairs(0)	sve-Drop(0)	sve-Late(0)	"
			. "sve-LongRR(0)	sve-LongRR_time	sve-Single(0)	sve-Bigem(0)	sve-Trigem(0)	sve-AF(0)"
	} else if (monType="BGH") {
		tabs := "dem-Name_L	dem-Name_F	dem-MRN	dem-Ordering	dem-Sex(NA)	dem-DOB	dem-VOID_Practice	dem-Indication	"
			. "dem-Test_date	dem-Test_end	dem-VOID	dem-Billing	"
			. "counts-Critical(0)	counts-Total(0)	counts-Serious(0)	counts-Manual(0)	counts-Stable(0)	counts-Auto(0)"
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

makeORU(wqid) {
/*	Generate a outbound ORU message for Epic
	Real world incoming Preventice ORU MSH.8 is a Preventice number.
	If MSH.8 contains "EPIC", was generated from MakeTestORU(),	so test ORU will set to OBR.32 and OBX.5 as "###" for filling in by Access DB
*/
	global fldval, hl7out, montype, isDevt
	dict:=readIni("EpicResult")
	ep:=readIni("epRead")
	
	hl7time := A_Now
	hl7out := Object()
	
	buildHL7("MSH"
		,{1:"^~\&"
		, 2:"CVTRRIQ"
		, 3:"CVTRRIQ"
		, 4:"HS"
		, 6:hl7time
		, 8:"ORU^R01"
		, 9:wqid
		, 10:"T"
		, 11:"2.5.1"})
	
	buildHL7("PID"
		,{2:fldval["dem-MRN"]
		, 3:fldval["dem-MRN"] "^^^^CHRMC"
		, 5:fldval["dem-Name_L"] "^" fldval["dem-Name_F"]
		, 7:parseDate(fldval["dem-DOB"]).YMD
		, 8:substr(fldval["dem-Sex"],1,1)
		, 18:fldval.accountnum})
	
	buildHL7("PV1"
		,{19:fldval.encnum
		, 50:wqid})
	

/*	Insert fake RTF and reading EP
	and monType in OBR_4 in cutover condition
*/
	if (isDevt=true) {
		MsgBox, 36, Testing, Create ORU with fake RTF and reading EP?
	}
	IfMsgBox, Yes
	{
	;~ if (fldval.MSH_ctrlID~="EPIC") {
		FileRead, rtf, .\files\test-RTF.txt
		EPdoc := ep[fldval["dem-reading"]]
	} 
	else
	{
		rtf := "###"
		EPdoc := "###"
	}
	obr4 := (montype~="i)PR|Hol") ? "CVCAR02^HOLTER MONITOR 24 HOUR^IMGEAP"
			: (montype~="i)BGH") ? "CVCAR05^CARDIAC EVENT MONITOR^IMGEAP"
			: (montype~="i)Mini|ZIO") ? "CVCAR102^HOLTER MONITOR 3-15 DAY^IMGEAP"
			: ""
	if (fldval.oldUID) {
		obr4 := RegExReplace(obr4,"^(.*?)\^","CVCAR104^")
		obr4 := RegExReplace(obr4,"\^IMGEAP$"," (cutover)^IMGEAP")
	}
	fldval.obr4 := obr4
	buildHL7("OBR"
		,{2:fldval.order
		, 3:fldval.accession
		, 4:fldval.obr4
		, 7:fldval.date
		, 16:fldval.OBR_ProviderCode "^"
			. fldval.OBR_ProviderNameL "^"
			. fldval.OBR_ProviderNameF
			. "^^^^^^MSOW_ORG_ID"
		, 25:"F"
		, 32:EPdoc })																	; Epic test: Substitute reading EP string "NPI^LAST^FIRST"
	
	buildHL7("OBX"
		,{2:"FT"
		, 3:"&GDT^HOLTER/EVENT RECORDER REPORT"
		, 5:rtf																			; Epic test: Substitute test rtf
		, 11:"F"
		, 14:hl7time})
	
	if (montype~="BGH") {																; no DDE for CEM
		return
	}
	
	for key,val in dict																	; Loop through all values in Dict (from ini)
	{
		str:=StrSplit(val,"^")
		buildHL7("OBX"																	; generate OBX for each value
			,{2:"TX"
			, 3:key "^" str.1 "^IMGLRR"
			, 5:fldval[str.2]
			, 11:"F"
			, 14:hl7time})
	}
	
	return
}

makeTestORU() {
/*	Generate a fake Preventice inbound ORU message based on the Preventice ORM registration data
*/
	global ptDem, hl7out, path
	hl7time := A_Now
	hl7out := Object()
	PVID := "2459720"
	
	buildHL7("MSH"
		,{1:"^~\&"
		, 2:"ADEPTIA"
		, 3:"ECARDIO"
		, 5:"8382"
		, 6:hl7time
		, 8:"ORU^R01"
		, 9:"EPIC" A_TickCount
		, 10:"P"
		, 11:"2.5"})
	
	buildHL7("PID"
		,{2:PVID
		, 3:ptDem.MRN
		, 5:ptDem.nameL "^" ptDem.nameF
		, 7:parseDate(ptDem.dob).YMD
		, 8:substr(ptDem.sex,1,1)
		, 11:ptDem.Addr1 "^" ptDem.Addr2 "^" ptDem.city "^" ptDem.state "^" ptDem.zip
		, 13:ptDem.phone
		, 18:PVID })
	
	tmpPrv := parseName(ptDem.provider)
	buildHL7("PV1"
		,{7:ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, 39:A_now })
	
	buildHL7("OBR"
		,{2:ptDem.wqid
		, 3:PVID
		, 4:strQ(ptDem.model~="Mortara" ? 1 : "","Holter^Holter")
			. strQ(ptDem.model~="Heart" ? 1 : "","CEM^CEM")
			. strQ(ptDem.model~="Mini" ? 1 : "","Holter^Holter")
		, 7:hl7time
		, 16:ptDem.NPI "^" tmpPrv.last "^" tmpPrv.first
		, 20:"OnComplete"
		, 22:A_now })
	
	buildHL7("OBX"
		,{2:"TX"
		, 3:strQ(ptDem.model~="Mortara" ? 1 : "","Holter^Holter")
			. strQ(ptDem.model~="Heart" ? 1 : "","CEM^CEM")
			. strQ(ptDem.model~="Mini" ? 1 : "","Holter^Holter")
		, 11:"F"
		, 14:A_now })
	
	FileRead, testTXT, % ".\files\test-ED_"
		. strQ(ptDem.model~="Mortara" ? 1 : "","HOL")
		. strQ(ptDem.model~="Heart" ? 1 : "","CEM")
		. strQ(ptDem.model~="Mini" ? 1 : "","MCT")
		. ".txt"
	buildHL7("OBX"
		,{2:"ED"
		, 3:"PDFReport1^PDF Report^^^^"
		, 4:ptDem.nameL "_" ptDem.nameF "_" ptDem.mrn "_" parseDate(ptDem.dob).YMD "_" A_now ".pdf"
		, 5:testTXT
		, 6:8
		, 11:"F"
		, 14:A_now })
	
	buildHL7("OBX",{2:"NM|Brady_AvgRate^Bradycardia average rate^Preventice^^^||51|bpm|||||" })
	buildHL7("OBX",{2:"TX|Brady_LongestDur^Bradycardia longest duration^Preventice^^^||01:09:06|time|||||" })
	buildHL7("OBX",{2:"DTM|Brady_LongestDur_Dt^Date and Time of longest Bradycardia episode^Preventice^^^||20191206012300|datetime|||||" })
	buildHL7("OBX",{2:"TX|Brady_ShortestDur^Bradycardia shortest duration^Preventice^^^||00:00:06|time|||||" })
	buildHL7("OBX",{2:"DTM|Brady_ShortestDur_Dt^Date and Time of shortest Bradycardia episode^Preventice^^^||20191115171500|datetime|||||" })
	buildHL7("OBX",{2:"TX|Diagnosis^Diagnosis (Indication for Monitoring)^Preventice^^^||R00.2: Palpitations||||||" })
	buildHL7("OBX",{2:"TX|Disconnect_Dur^Overall disconnect duration^Preventice^^^||3.18:26:04|time|||||" })
	buildHL7("OBX",{2:"DTM|Enroll_End_Dt^Enrollment End Date^Preventice^^^||20191206000000|datetime|||||" })
	buildHL7("OBX",{2:"DTM|Enroll_Start_Dt^Enrollment Start Date^Preventice^^^||20191107000000|datetime|||||" })
	buildHL7("OBX",{2:"NM|HTRate_MaxRate^Maximum heart rate^Preventice^^^||162|bpm|||||" })
	buildHL7("OBX",{2:"NM|HTRate_MeanRate^Mean heart rate^Preventice^^^||69|bpm|||||" })
	buildHL7("OBX",{2:"NM|HTRate_MinRate^Minimum heart rate^Preventice^^^||38|bpm|||||" })
	buildHL7("OBX",{2:"NM|Pause_Count^Pauses >= 3 seconds^Preventice^^^||0||||||" })
	
	FileAppend
		, % hl7out.msg
		, % path.PrevHL7in ptDem.nameL "_" ptDem.nameF "_" ptDem.mrn "_" parseDate(ptDem.dob).YMD "_" A_now ".hl7"
	
	return
}

shortenPDF(find) {
	eventlog("ShortenPDF")
	global fileIn, fileNam, wincons
	sleep 500
	fullNam := filenam "full.txt"

	Progress,,% " ",Scanning full size PDF...
	RunWait, .\files\pdftotext.exe "%fileIn%" "%fullnam%",,min,wincons					; convert PDF all pages to txt file
	eventlog("Extracting full text.")
	progress,100,, Shrinking PDF...
	FileRead, fulltxt, %fullnam%
	findpos := RegExMatch(fulltxt,find)
	pgpos := instr(fulltxt,"Page ",,findpos-strlen(fulltxt))
	RegExMatch(fulltxt,"Oi)Page\s+(\d+)\s",pgs,pgpos)
	pgpos := pgs.value(1)
	RunWait, .\files\pdftk.exe "%fileIn%" cat 1-%pgpos% output "%fileIn%-sh.pdf",,min
	if !FileExist(fileIn "-sh.pdf") {
		FileCopy, %fileIn%, %fileIn%-sh.pdf
	}
	filedelete, %fullnam%
	FileGetSize, sizeIn, %fileIn%
	FileGetSize, sizeOut, %fileIn%-sh.pdf
	eventlog("IN: " thousandsSep(sizeIn) ", OUT: " thousandsSep(sizeOut))
	progress, off
return	
}

findFullPdf(wqid:="") {
/*	Scans HolterDir for potential full disclosure PDFs
	maybe rename if appropriate
*/
	global path, fldval, pdfList, AllowSavedPDF
	
	pdfList := Object()																	; clear list to add to WQlist
	pdfScanPages := 3
	
	fileCount := ComObjCreate("Scripting.FileSystemObject").GetFolder(path.holterPDF).Files.Count
	
	Loop, files, % path.holterPDF "*.pdf"
	{
		fileIn := A_LoopFileFullPath													; full path and filename
		fname := A_LoopFileName															; full filename
		fnam := RegExReplace(fname,"i)\.pdf")											; filename without ext
		progress, % 100*A_index/fileCount, % fname, Scanning PDFs folder
		
		;---Skip any PDFs that have already been processed or are in the middle of being processed
		if (fname~="i)(-sh|-short)\.pdf") 
			continue
		if FileExist(fname "-sh.pdf") 
			continue
		if FileExist(fnam "-short.pdf") 
			continue
		
		RegExMatch(fname,"O)_WQ([A-Z0-9]+)(_\w)?\.pdf",fnID)									; get filename WQID if PDF has already been renamed
		
		if (readWQ(fnID.1).node = "done") {
			eventlog("Leftover PDF: " fnam ", moved to archive.")
			FileMove, % fileIn, % path.holterPDF "archive\" fname, 1
			continue
		}
		
		if (fnID.0 = "") {																; Unmatched full disclosure PDF
			RunWait, .\files\pdftotext.exe -l %pdfScanPages% "%fileIn%" "%fnam%.txt",,min		; convert PDF pages with no tabular structure
			FileRead, newtxt, %fnam%.txt												; load into newtxt
			FileDelete, %fnam%.txt
			StringReplace, newtxt, newtxt, `r`n`r`n, `r`n, All							; remove double CRLF
			
			flds := getPdfID(newtxt)
			
			if (AllowSavedPDF="true") && InStr(flds.wqid,"00000") {
				eventlog("Unmatched PDF: " fileIn)
				continue
			}
			if (AllowSavedPDF!="true") && (flds.type = "E") {
				MsgBox, 262160, File error
					, % path.holterPDF "`n" fName "`n"
					. "saved from email.`n`n"
					. "DO NOT SAVE FROM EMAIL!`n`n"
					. "(delete the file to stop getting this message)"
				eventlog("CEM saved from email: " fileIn)
				continue
			}
			
			newFnam := strQ(flds.nameL,"###_" flds.mrn,fnam) strQ(flds.wqid,"_WQ###")
			FileMove, %fileIn%, % path.holterPDF newFnam ".pdf", 1						; rename the unprocessed PDF
			eventlog("Holter PDF: " fName " renamed to " newFnam)
			fName := newFnam ".pdf"
		} 
		if !objhasvalue(pdfList,fName) {
			pdfList.push(fName)
		}
		
		if (wqid = "") {																; this is just a refresh loop
			continue																	; just build the list
		}
		
		if (fnID.1 == wqid) {															; filename WQID matches wqid arg
			FileMove, % path.PrevHL7in fldval.Filename, % path.PrevHL7in fldval.Filename "-sh.pdf"		; rename the pdf in hl7dir to -short.pdf
			FileMove, % path.holterPDF fName , % path.PrevHL7in fldval.filename 		; move this full disclosure PDF into hl7dir
			progress, off
			eventlog(fName " moved to " path.PrevHL7in)
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
			res.date := dt.YMD
			res.time := dt.hr dt.min
		dobDt := parseDate(trim(stregX(txt,"(Date of Birth|DOB):?",1,1,"\R",1)))
			res.dob := dobDt.YMD
		res.mrn := trim(stregX(txt,"Secondary ID:?",1,1,"Age:?",1))
		res.ser := trim(stregX(txt,"Recorder (No|Number):?",1,1,"\R",1))
		res.wqid := strQ(findWQid(res.date,res.mrn,"Mortara H3+ - " res.ser).id,"###","00000") "_H"
	} else if instr(txt,"BodyGuardian Heart") {											; BG Heart
		res.type := "E"
		name := parseName(res.name := trim(stregX(txt,"Patient:",1,1,"Enrollment Info|Patient ID",1)," `t`r`n"))
			res.nameL := name.last
			res.nameF := name.first
		dt := parseDate(trim(stregX(txt,"Period \(.*?\R",1,1," - ",1)," `t`r`n"))
			res.date := dt.YMD
		res.mrn := trim(stregX(txt,"Patient ID",1,1,"Gender",1)," `t`r`n")
		res.wqid := strQ(findWQid(res.date,res.mrn).id,"###","00000") "_E"
	} else if instr(txt,"Zio XT") {														; Zio
		res.type := "Z"
		name := parseName(res.name := trim(stregX(txt,"Final Report for\R",1,1,"\R",1)," `t`r`n"))
			res.nameL := name.last
			res.nameF := name.first
		enroll := stregX(txt,"Enrollment Period",1,0,"Analysis Time",1)
		dt := parseDate(stregX(enroll,"i)\R+.*?(hours|days).*?\R+",1,1,",",1))
			res.date := dt.YMD
		res.mrn := strQ(trim(stregX(txt,"Patient ID\R",1,1,"\R",1)," `t`r`n"),"###","Zio")
		res.wqid := "00000_Z"
	} else if instr(txt,"Preventice Services, LLC") {									; BG Mini report
		res.type := "M"
		name := parseName(res.name := trim(stregX(txt,"Patient Name:",1,1,"\R",1)))
			res.nameL := name.last
			res.nameF := name.first
		dt := parseDate(trim(stregX(txt,"Test Start:",1,1,"Test End:",1)))
			res.date := dt.YMD
			res.time := dt.hr dt.min
		dobDt := parseDate(trim(stregX(txt,"(Date of Birth|DOB):",1,1,"\R",1)))
			res.dob := dobDt.YMD
		res.mrn := trim(stregX(txt,"MRN:",1,1,"Date of Birth:",1)," `r`n")
		res.ser := trim(stregX(txt,"Device Serial Number:",1,1,"\(Firmware",1))
		res.wqid := strQ(findWQid(res.date,res.mrn).id,"###","00000") "_M"
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
	eventlog("CheckProc")
	fetchQuit := false
	
	if !(fldval.wqid) {
		id := findWQid(parseDate(fldval["dem-Test_date"]).YMD							; search wqid based on combination of study date, mrn, SN
				, fldval["dem-MRN"]
				, fldval["dem-Device_SN"]).id
		if (id) {																		; pull some vals
			res := readWQ(id)
			fldval["dem-Device_SN"] := strX(res.dev," ",0,1,"",0)
			fldval.name := res.name
			fldval.node := res.node
			fldval.wqid := id
			eventlog("CheckProc: found wqid " id " in " res.node)
		} else {
			eventlog("CheckProc: no matching wqid found")
		}
	}
	if (fldval.node = "done") {
	;~ if (zzzfldval.node = "done") {
		MsgBox % fileIn " has been scanned already.`n`nDeleting file."
		eventlog(fileIn " already scanned. PDF deleted.")
		FileDelete, % fileIn
		fetchQuit := true
		return
	}
	
	ptDem := Object()																	; Populate temp object ptDem with parsed data from HL7 or PDF fldVal
	ptDem["nameL"] := fldVal["dem-Name_L"]												; dem-Name contains ['] not [^]
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
			. "Paste clipboard into Epic search to select patient and encounter"
		
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
		fldVal["dem-Name_L"] := fldval["Name_L"] := RegExReplace(ptDem["nameL"],"\^","'")
		fldVal["dem-Name_F"] := fldval["Name_F"] := RegExReplace(ptDem["nameF"],"\^","'")
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
				id := makeUID() 															; create wqid record if it doesn't exist somehow
				wq.addElement("enroll","/root/pending",{id:id})
				fldval.wqid := id
			}
			newID := "/root/pending/enroll[@id='" id "']"
			ptDem.date := parseDate(ptDem["EncDate"]).YMD
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
	fldVal["dem-Name_L"] := fldval["Name_L"] := RegExReplace(ptDem["nameL"],"\^","'")
	fldVal["dem-Name_F"] := fldval["Name_F"] := RegExReplace(ptDem["nameF"],"\^","'")
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

Holter_BGM_HL7:
{
	eventlog("Holter_BGMini_HL7")
	monType := "Mini"
	
	fldval["dem-Test_date"] := parsedate(fldval["Enroll_Start_Dt"]).MDY
	fldval["dem-Test_end"]	:= parsedate(fldval["Enroll_End_Dt"]).MDY
	fldval["dem-Recording_time"] := strQ(fldval["Monitoring_Period"], parsedate("###").DHM
									, calcDuration(fldval["hrd-Total_Time"]).DHM " (DD:HH:MM)")
	fldval["dem-Analysis_time"] := strQ(fldval["Analyzed_Data"], parsedate("###").DHM
									, calcDuration(fldval["hrd-Analyzed_Time"]).DHM " (DD:HH:MM)")

	gosub checkProc												; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	fieldsToCSV()
	fieldcoladd("","INTERP","")									; fldval["Narrative"]
	fieldcoladd("","Mon_type","Holter")
	
	FileCopy, %fileIn%, %fileIn%-sh.pdf
	
	fldval.done := true

return
}

Holter_BGM:
{
	eventlog("Holter_BGMini")
	monType := "Mini"
	
	/* Pulls text between field[n] and field[n+1], place in labels[n] name, with prefix "dem-" etc.
	 */
	demog := columns(newtxt,"Patient\s+Information","Ventricular Tachycardia",1,"Test Start")
	;~ fields[1] := ["MRN","Patient Name","Age","Date Of Birth","Gender","Site"
				;~ , "Test Start","Test End","Test Duration","Analysis Duration"]
	;~ labels[1] := ["MRN","Name","null","DOB","Sex","null"
				;~ , "Test_date","Test_end","Recording_time","Analysis_time"]
	fields[1] := ["Test Start","Test End","Test Duration","Analysis Duration"]
	labels[1] := ["Test_date","Test_end","Recording_time","Analysis_time"]
	scanParams(demog,1,"dem",1)
	
	t0 := parseDate(fldval["dem-Test_date"]).ymd
	;~ t1 := t0.YMD t0.hr t0.min t0.sec
	
	summary := columns(newtxt,"\s+Ventricular Tachycardia","\s+Interpretation",,"Total QRS") "<<<end"
	daycount(summary,t0)
	
	sumEvent := stregX(summary,"",1,0,"\s+Summary\R",1) "<<<end"
	summary := stregX(summary,"\s+Summary\R",1,1,"<<<end",0)
	
	sumTot := stregX(summary,"\s+Totals\R",1,1,"\s+Heart Rate\R",1)
	
	sumRate := sumTot "`n" stregX(summary,"\s+Heart Rate\R",1,1,"\s+Ventricular Event Information\R",1)
	fields[1] := ["Total QRS","Minimum","Maximum","Average","Tachycardia","Bradycardia"]
	labels[1] := ["Total_beats","Min","Max","Avg","Longest_tachy","Longest_brady"]
	scanParams(sumRate,1,"hrd",1)
	
	sumVE := sumTot "`n" stregX(summary,"\s+Ventricular Event Information\R",1,1,"\s+Supraventricular Event Information\R",1)
	fields[2] := ["Ventricular","Isolated","Bigeminy","Couplets","Total Runs","Longest","Fastest"]
	labels[2] := ["Total","SingleVE","Bigeminy","Couplets","Runs","Longest","Fastest"]
	scanParams(sumVE,2,"ve",1)
	
	sumSVE := sumTot "`n" stregX(summary,"\s+Supraventricular Event Information\R",1,1,"\s+RR.Pause\R",1)
	fields[3] := ["Supraventricular","Isolated","Couplets","Total Runs","Longest","Fastest"]
	labels[3] := ["Total","Single","Pairs","Runs","Longest","Fastest"]
	scanParams(sumSVE,3,"sve",1)
	
	sumPause := stregX(summary,"\s+RR.Pause\R",1,1,"\s+AFib.AFlutter\R",1)
	fields[4] := ["Maximum","Total Pauses"]
	labels[4] := ["LongRR","Pauses"]
	scanParams(sumPause,4,"sve",1)
	
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

daycount(byref txt,day1) {
	n:="(\d{2}:\d{2}:\d{2}) Day (\d{1,2})"
	pos:=1, v:=0
	while pos:=RegExMatch(txt,n,m,v+pos)
	{
		day:=day1
		day += m2, Days
		v:=StrLen(m)
		txt:=RegExReplace(txt,n,parseDate(substr(day,1,8)).mdy " at " m1,,1,pos)
	}
	return
}

Zio:
{
	eventlog("Holter_Zio")
	monType := "Zio"
	
	RunWait, .\files\pdftotext.exe -table -fixed 3 "%fileIn%" "%filenam%.txt", , hide			; reconvert entire Zio PDF 
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
	
	FileCopy, %fileIn%, %fileIn%-sh.pdf
	
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
	
	if !(obxVal["Enroll_Start_Dt"]) {													; missing this if no OBX
		eventlog("No OBX data.")
		gosub processPDF																; process as an ad hoc
		return																			; and bail out
	}
	
	fieldcoladd("dem","Test_date",niceDate(obxVal["Enroll_Start_Dt"]))
	fieldcoladd("dem","Test_end",niceDate(obxVal["Enroll_End_Dt"]))
	
	count_block := stregX(newtxt,"Event Counts",1,1,"Summarized|Rhythm",1)
	count_block := RegExReplace(count_block,"(\d) ","$1`r`n")
	fields[3] := ["Critical","Total","Serious","(Manual|Pt Trigger)","Stable","Auto Trigger","\R"]
	labels[3] := ["Critical","Total","Serious","Manual","Stable","Auto","null"]
	
	if (fldval["counts-Auto"]="" && fldval["counts-Manual"]="")							; No Event Counts values
	{																					; parse from PDF
		fieldvals(count_block,3,"counts")
	} 
	else																				; Still no Event Counts (bad PDF)
	{
		count:=[]																		; create object for counts
		count["Patient-Activated"]:=0													; zero the results instead of null
		count["Auto-Detected"]:=0
		count["Stable"]:=0
		count["Serious"]:=0
		count["Critical"]:=0
		for key,val in obxVal															; recurse through obxVal results
		{
			if (key~="Event_Acuity|Event_Type") {										; count Critical/Serious/Stable and Auto/Manual events
				count[val] ++															; more reliable than parsing PDF
			}
		}
		fieldcoladd("counts","Critical",count["Critical"])
		fieldcoladd("counts","Serious",count["Serious"])
		fieldcoladd("counts","Stable",count["Stable"])
		fieldcoladd("counts","Manual",count["Patient-Activated"])
		fieldcoladd("counts","Auto",count["Auto-Detected"])
		fieldcoladd("counts","Total",count["Auto-Detected"]+count["Patient-Activated"])
		eventlog("Event Count block not parsed, counted from OBR.")
	}
	
	gosub checkProc												; check validity of PDF, make demographics valid if not
	if (fetchQuit=true) {
		return													; fetchGUI was quit, so skip processing
	}
	
	fieldstoCSV()
	
	fieldcoladd("","Mon_type","Event")
	
	FileCopy, %fileIn%, %fileIn%-sh.pdf
	
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
	
	FileCopy, %fileIn%, %fileIn%-sh.pdf
	
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
	StringReplace, x, x, `r`n, `n, all
	
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

	if RegExMatch(txt,"(\d{1,2}) hr (\d{1,2}) min",t) {						; convert "24 hr 0 min" to "24:00"
		txt := t1 ":" zDigit(t2)
	}
	txt:=RegExReplace(txt,"i)( BPM)|( Event(s)?)|( Beat(s)?)|( sec(ond)?(s)?)")		; Remove units from numbers
	txt:=RegExReplace(txt,"(:\d{2}?)(AM|PM)","$1 $2")						; Fix time strings without space before AM|PM
	txt:=RegExReplace(txt,"\(DD:HH:MM:SS\)")								; Remove time units
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
		txt := parseDate(txt).mdy
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

;	Body Guardian Mini specific fixes
	if (monType="Mini") {
		if (lab ~= "Test_(date|end)") {													; convert dates to MDY format
			txt := parseDate(txt).mdy
		}
		if RegExMatch(txt																; reconstitute Beats and BPM for longest/fastest/slowest fields
		,"(.*)? \((\d{1,2}/\d{1,2}/\d{2,4} at \d{1,2}:\d{2}:\d{2})\)"
		,res) {
			res1 := RegExReplace(res1,"(\d+)\s*,\s*(\d+)","$1 beats, $2 bpm")
			fieldColAdd(pre,lab,res1)
			fieldColAdd(pre,lab "_time",res2)
			return
		}
		if RegExMatch(txt,"(.*?)\((.*?%)\)",res) {
			fieldColAdd(pre,lab,res1)
			fieldColAdd(pre,lab "_per",res2)
			return
		}
		if (lab~="_time") {
			if RegExMatch(txt,"(\d{1,2}):(\d{2}):\d{2}:\d{2}",res) {					; convert DD:HH:MM:SS into Days & Hrs
				txt := res1 " days, " res2 " hours"
			}
			if (txt~="^\d{14}$") {														; yyyymmddhhmmss
				txt := parseDate(txt).DT												; = mm/dd/yyyy at hh:mm:ss
			}
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
	fullNameDem := RegExReplace(ptDem["nameL"] ", " ptDem["nameF"],"\^","'")			; fetched ptDem.nameL has [^], passed nameL has [']
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

IEopen() {
/*	Use ComObj to open IE
	If not open, create a new instance
	If IE open, choose that windows object
	Return the IE window object
*/
	if !winExist("ahk_exe iexplore.exe") {
		wb := ComObjCreate("InternetExplorer.application")
		wb.visible := false
		return wb
	} 
	else {
		for wb in ComObjCreate("Shell.Application").Windows() {
			if InStr(wb.FullName, "iexplore.exe") {
				return wb
			}
		}
	}
}

IEclose() {
	DetectHiddenWindows, On
	while WinExist("ahk_class IEFrame")
	{
		i := A_index
		Process, Close, iexplore.exe
		sleep 500
	}
	
	return
}

httpComm(verb) {
	url := "https://depts.washington.edu/pedcards/change/direct.php?" 
			. "do=" . verb
	
	whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")							; initialize http request in object whr
	whr.Open("GET"																; set the http verb to GET file "change"
		, url
		, true)
	whr.Send()																	; SEND the command to the address
	whr.WaitForResponse()														; and wait for the http response
	return whr.ResponseText
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
			if (aValue="") {															; null aValue in "RX" is error
				return, false, errorlevel := 1
			}
			if (val ~= aValue) {														; val=text, aValue=RX
				return, key, Errorlevel := 0
			}
			if (aValue ~= val) {														; aValue=text, val=RX
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
	global user, userinstance
	comp := A_ComputerName
	FormatTime, sessdate, A_Now, yyyy.MM
	FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
	name := "logs/" . sessdate . ".log"
	txt := now " [" user "/" comp "/" userinstance "] " event "`n"
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
	if (x="") {
		return error
	}
	x := trim(x)																		; trim edges
	; x := RegExReplace(x,"\'","^")														; replace ['] with [^] to avoid XPATH errors
	x := RegExReplace(x," \w "," ")														; remove middle initial: Troy A Johnson => Troy Johnson
	x := RegExReplace(x,"(,.*?)( \w)$","$1")											; remove trailing MI: Johnston, Troy A => Johnston, Troy
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
	else if (ct>1)																		; James Jacob Jingleheimer Schmidt
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
	
	return {first:first
			, last:last
			, firstlast:first " " last
			, lastfirst:last ", " first 
			, init:substr(first,1,1) substr(last,1,1) }
}

ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-\._/]"
	date := []
	time := []
	x := RegExReplace(x,"[,\(\)]")
	
	if (x~="\d{4}.\d{2}.\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z") {
		x := RegExReplace(x,"[TZ]","|")
	}
	if (x~="\d{4}.\d{2}.\d{2}T\d{2,}") {
		x := RegExReplace(x,"T","|")
	}
	
	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",d) {			; 03-Jan-2015
		date.dd := zdigit(d1)
		date.mmm := d2
		date.mm := zdigit(objhasvalue(mo,d2))
		date.yyyy := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"\b(\d{4})[\-\.](\d{2})[\-\.](\d{2})\b",d) {					; 2015-01-03
		date.yyyy := d1
		date.mm := d2
		date.mmm := mo[d2]
		date.dd := d3
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
	else if RegExMatch(x,"i)(" moStr ")\s+(\d{1,2}),?\s+(\d{4})",d) {					; Dec 21, 2018
		date.mmm := d1
		date.mm := zdigit(objhasvalue(mo,d1))
		date.dd := zdigit(d2)
		date.yyyy := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"\b(19\d{2}|20\d{2})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",d)  {	; 20150103174307 or 20150103
		date.yyyy := d1
		date.mm := d2
		date.mmm := mo[d2]
		date.dd := d3
		date.date := d1 "-" d2 "-" d3
		
		time.hr := d5
		time.min := d6
		time.sec := d7
		time.time := d5 ":" d6 . strQ(d7,":###")
	}
	
	if RegExMatch(x,"iO)(\d+):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",t) {				; 17:42 PM
		hasDays := (t.value[4]) ? true : false 											; 4 nums has days
		time.days := (hasDays) ? t.value[1] : ""
		time.hr := trim(t.value[1+hasDays])
		if (time.hr>23) {
			time.days := floor(time.hr/24)
			time.hr := mod(time.hr,24)
			DHM:=true
		}
		time.min := trim(t.value[2+hasDays]," :")
		time.sec := trim(t.value[3+hasDays]," :")
		time.ampm := trim(t.value[5])
		time.time := trim(t.value)
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, days:zdigit(time.days)
			, hr:zdigit(time.hr), min:zdigit(time.min), sec:zdigit(time.sec)
			, ampm:time.ampm, time:time.time
			, DHM:zdigit(time.days) ":" zdigit(time.hr) ":" zdigit(time.min) " (DD:HH:MM)" 
 			, DT:date.mm "/" date.dd "/" date.yyyy " at " zdigit(time.hr) ":" zdigit(time.min) ":" zdigit(time.sec) }
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

; Convert duration secs to DDHHMMSS
calcDuration(sec) {
	DD := divTime(sec,"D")
	HH := divTime(DD.rem,"H")
	MM := divTime(HH.rem,"M")
	SS := MM.rem

	return { DHM: zDigit(DD.val) ":" zDigit(HH.val) ":" zDigit(MM.val)
			, DHMS: zDigit(DD.val) ":" zDigit(HH.val) ":" zDigit(MM.val) ":" zDigit(SS.val) }
}

divTime(sec,div) {
	static T:={D:86400,H:3600,M:60,S:1}
	xx := Floor(sec/T[div])
	rem := sec-xx*T[div]
	Return {val:xx,rem:rem}
}

ThousandsSep(x, s=",") {
; from https://autohotkey.com/board/topic/50019-add-thousands-separator/
	return RegExReplace(x, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" s)
}

ToBase(n,b) {
/*	from https://autohotkey.com/board/topic/15951-base-10-to-base-36-conversion/
	n >= 0, 1 < b <= 36
*/
   Return (n < b ? "" : ToBase(n//b,b)) . ((d:=mod(n,b)) < 10 ? d : Chr(d+55))
}

formatPhone(txt) {
; format phone as aaa-bbb-cccc
	return RegExReplace(txt,".*?(\d{3}).*?(\d{3}).*?(\d{4})","$1-$2-$3")
}

WriteOut(parentpath,node) {
	global wq
	
	filecheck()
	FileOpen(".lock", "W")																; Create lock file.
	locPath := wq.selectSingleNode(parentpath)
	locNode := locPath.selectSingleNode(node)
	clone := locNode.cloneNode(true)													; make copy of wq.node
	
	if !IsObject(locNode) {
		eventlog("No such node <" parentpath "/" node "> for WriteOut.")
		FileDelete, .lock																; release lock file.
		return error
	}
	
	z := new XML("worklist.xml")														; load a copy into z
	
	if !IsObject(z.selectSingleNode(parentpath "/" node)) {								; no such node in z
		z.addElement("newnode",parentpath)												; create a blank node
		node := "newnode"
	}
	zPath := z.selectSingleNode(parentpath)												; find same "node" in z
	zNode := zPath.selectSingleNode(node)
	zPath.replaceChild(clone,zNode)														; replace existing zNode with node clone
	
	writeSave(z)
	
	FileDelete, .lock
	
	return
}

WriteSave(z) {
/*	Saves worklist.xml with integrity check
	presence of .lock does not matter
*/
	global wq
	
	loop, 3
	{
		z.transformXML()
		z.save("worklist.xml")
		FileRead,wltxt,worklist.xml
		
		if instr(substr(wltxt,-9),"</root>") {
			valid:=true
			break
		}
		
		eventlog("WriteSave failed " A_index)
		sleep 2000
	}
	
	if (valid=true) {
		FileCopy, worklist.xml, bak\%A_now%.bak
		wq := z
	}
	
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
	IniRead,x,.\files\trriq.ini,%section%
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

#Include CMsgBox-img.ahk
#Include InputBox.ahk
#Include Class_LV_Colors.ahk
#Include xml.ahk
#Include sift3.ahk
#Include hl7.ahk
#Include callSched.ahk
