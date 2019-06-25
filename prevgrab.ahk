/*	PrevGrab - Grabs data from Preventice website
	Saves results for TRRIQ
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force  ; only allow one running instance per user
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

		if (phase="Enrollment") {
			eventlog("Update Preventice enrollments.")
			CheckPreventiceWeb("Patient Enrollment")
		}
		if (phase="Inventory") {
			eventlog("Update Preventice inventory.")
			CheckPreventiceWeb("Facilities")
		}

ExitApp

CheckPreventiceWeb(win) {
	global phase
	SetTimer, idleTimer, Off
	
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
		if (clip=clip0) {																; no change since last clip
			progress, off
			MsgBox,4144,, Reached the end of novel records.`n`n%phase% update complete!
			break
		}
		
		done := %prvFunc%(tbl)		; parsePreventiceEnrollment() or parsePreventiceInventory()
		
		if (done=0) {																	; no new records returned
			progress, off
			MsgBox,4144,, Reached the end of novel records.`n`n%phase% update complete!
			break
		}
		clip0 := clip																	; set the check for repeat copy
		
		PreventiceWebPager(wb,str[phase].changed,str[phase].btn)
	}
	
	setwqupdate()
	
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
	checkdays := 21
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
		res.name := parsename(res.name).lastfirst
		date := parseDate(res.date).YMD
		dt := A_Now
		dt -= date, Days
		if (dt>checkdays) {																; if days > threshold, break loop
			break
		} else {																		; otherwise done+1 == keep paging
			done ++
		}
		
	/*	Check whether any params match this device
	*/
		if enrollcheck("[mrn='" res.mrn "'][date='" date "'][dev='" res.dev "']") {		; MRN+DATE+S/N = perfect match
			continue
		}
		if (id:=enrollcheck("[mrn='" res.mrn "'][dev='" res.dev "']")) {				; MRN+S/N, no DATE
			en:=readWQ(id)
			if (en.node="done") {
				continue
			}
			wqSetVal(id,"date",date)
			eventlog(en.name " (" id ") changed WQ date '" en.date "' ==> '" date "'")
			continue
		}
		if (id:=enrollcheck("[mrn='" res.mrn "'][date='" date "']")) {					; MRN+DATE, no S/N
			en:=readWQ(id)
			if (en.node="done") {
				continue
			}
			wqSetVal(id,"dev",res.dev)
			eventlog(en.name " (" id ") changed WQ dev '" en.dev "' ==> '" res.dev "'")
			continue
		}
		if (id:=enrollcheck("[date='" date "'][dev='" res.dev "']")) {					; DATE+S/N, no MRN
			en:=readWQ(id)
			if (en.node="done") {
				continue
			}
			wqSetVal(id,"mrn",res.mrn)
			eventlog(en.name " (" id ") changed WQ mrn '" en.mrn "' ==> '" res.mrn "'")
			continue
		} 
		
	/*	No match (i.e. unique record)
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

enrollcheck(params) {
	global wq
	
	en := wq.selectSingleNode("//enroll" params)
	id := en.getAttribute("id")
	
; 	returns id if finds a match, else null
	return id																			
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

