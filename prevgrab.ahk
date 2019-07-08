/*	PrevGrab - Grabs data from Preventice website
	Saves results for TRRIQ
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force  ; only allow one running instance per user
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

Config: 
{
/*	Set config vals for script
*/
	global webstr:={}
		,  gl:={}
	
	webStr.Enrollment := {dlg:"Enrollment / Submitted Patients"
		, url:"https://secure.preventice.com/Enrollments/EnrollPatients.aspx?step=2"
		, win:"Patient Enrollment"
		, tbl:"ctl00_mainContent_PatientListSubmittedCtrl1_RadGridPatients_ctl00"
		, changed:"ctl00_mainContent_PatientListSubmittedCtrl1_lblTotalCountMessage"
		, btn:"ctl00_mainContent_PatientListSubmittedCtrl1_btnNextPage"
		, click:"getElementById(btnStr).click()"
		, fx:"ParsePreventiceEnrollment"}
	webStr.Inventory := {dlg:"Facility`nInventory Status`nDevice in Hand (Enrollment not linked)"
		, url:"https://secure.preventice.com/Facilities/"
		, win:"Facilities"
		, tbl:"ctl00_mainContent_InventoryStatus_userControl_gvInventoryStatus_ctl00"
		, changed:"ctl00_mainContent_InventoryStatus_userControl_gvInventoryStatus_ctl00_Pager"
		, btn:"rgPageNext"
		, click:"getElementsByClassName(btnStr)[0].click()"
		, fx:"ParsePreventiceInventory"}
	
	gl.TRRIQ_path := "\\childrens\files\HCCardiologyFiles\EP\Holter DB\TRRIQ"
	
	IfInString, A_ScriptDir, AhkProjects 
	{
		gl.isAdmin := true
		gl.files_dir := A_ScriptDir "\files"
		gl.user_name := "test"
		gl.user_pass := "test"
	} else {
		gl.isAdmin := false
		gl.files_dir := TRRIQ_path "\files"
		gl.user_name := "test"
		gl.user_pass := "test"
	}
}

MainLoop:
{
	eventlog("Update Preventice enrollments.")
	PreventiceWebGrab("Enrollment")
	
	eventlog("Update Preventice inventory.")
	PreventiceWebGrab("Inventory")
	
	ExitApp
}

PreventiceWebGrab(phase) {
	global webStr
	web := webStr[phase]
	
	wb := IEopen()
	wb.visible := true
	wb.Navigate(web.url)
	while wb.busy {
		sleep 10
	}
	
	;~ wb := ieGet(webStr[phase].win)
	;~ SetTimer, idleTimer, Off
	MsgBox end
	ComObjConnect(wb)
	WinKill, ahk_exe iexplore.exe
	ExitApp
	
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
	
	;~ setwqupdate()
	
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
	;~ fileCheck()
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
		;~ if enrollcheck("[mrn='" res.mrn "'][date='" date "'][dev='" res.dev "']") {		; MRN+DATE+S/N = perfect match
			;~ continue
		;~ }
		;~ if (id:=enrollcheck("[mrn='" res.mrn "'][dev='" res.dev "']")) {				; MRN+S/N, no DATE
			;~ en:=readWQ(id)
			;~ if (en.node="done") {
				;~ continue
			;~ }
			;~ wqSetVal(id,"date",date)
			;~ eventlog(en.name " (" id ") changed WQ date '" en.date "' ==> '" date "'")
			;~ continue
		;~ }
		;~ if (id:=enrollcheck("[mrn='" res.mrn "'][date='" date "']")) {					; MRN+DATE, no S/N
			;~ en:=readWQ(id)
			;~ if (en.node="done") {
				;~ continue
			;~ }
			;~ wqSetVal(id,"dev",res.dev)
			;~ eventlog(en.name " (" id ") changed WQ dev '" en.dev "' ==> '" res.dev "'")
			;~ continue
		;~ }
		;~ if (id:=enrollcheck("[date='" date "'][dev='" res.dev "']")) {					; DATE+S/N, no MRN
			;~ en:=readWQ(id)
			;~ if (en.node="done") {
				;~ continue
			;~ }
			;~ wqSetVal(id,"mrn",res.mrn)
			;~ eventlog(en.name " (" id ") changed WQ mrn '" en.mrn "' ==> '" res.mrn "'")
			;~ continue
		;~ } 
		
	/*	No match (i.e. unique record)
	 *	add new record to PENDING
	 */
		;~ sleep 1																			; delay 1ms to ensure different tick time
		;~ id := A_TickCount 
		;~ newID := "/root/pending/enroll[@id='" id "']"
		;~ wq.addElement("enroll","/root/pending",{id:id})
		;~ wq.addElement("date",newID,date)
		;~ wq.addElement("name",newID,res.name)
		;~ wq.addElement("mrn",newID,res.mrn)
		;~ wq.addElement("dev",newID,res.dev)
		;~ wq.addElement("prov",newID,filterProv(res.prov).name)
		;~ wq.addElement("site",newID,filterProv(res.prov).site)
		;~ wq.addElement("webgrab",newID,A_now)
		;~ done ++
		
		;~ eventlog("Added new registration " res.mrn " " res.name " " date ".")
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
	
	;~ writeout("/root","inventory")
	
	return true
}

IEopen() {
/*	Use ComObj to open IE
	If not open, create a new instance
	If IE open, choose that windows object
	Return the IE window object
*/
	if !winExist("ahk_exe iexplore.exe") {
		wb := ComObjCreate("InternetExplorer.application")
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

IEurl(url) {
/*	Open a URL
*/
	global wb
	
	wb.Navigate(url)																	; load URL
	while wb.busy {																		; wait until done loading
		sleep 10
	}
	
	if instr(wb.LocationURL,"UserLogin") {
		preventiceLogin()
	}
	
	return
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

ParseName(x) {
/*	Determine first and last name
*/
	if (x="") {
		return error
	}
	x := trim(x)																		; trim edges
	x := RegExReplace(x,"\'","^")														; replace ['] with [^] to avoid XPATH errors
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
		;~ last := cmsgbox("Name check",x "`n" RegExReplace(x,".","--") "`nWhat is the patient's`nLAST NAME?",q)
		if (last~="close|xClose") {
			return {first:"",last:x}
		}
		first := RegExReplace(x," " last)
	}
	
	return {first:first
			,last:last
			,firstlast:first " " last
			,lastfirst:last ", " first
			,apostr:RegExReplace(x,"\^","'")}
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
	else if RegExMatch(x,"\b(\d{4})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",d)  {			; 20150103174307
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
	
	if RegExMatch(x,"iO)(\d{1,2}):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",t) {				; 17:42 PM
		hasDays := (t.value[4]) ? true : false 												; 4 nums has days
		time.days := (hasDays) ? t.value[1] : ""
		time.hr := zdigit(t.value[1+hasDays])
		time.min := trim(t.value[2+hasDays]," :")
		time.sec := trim(t.value[3+hasDays]," :")
		time.ampm := trim(t.value[5])
		time.time := trim(t.value)
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, days:time.days, hr:time.hr, min:time.min, sec:time.sec, ampm:time.ampm, time:time.time}
}

zDigit(x) {
; Add leading zero to a number
	return SubStr("0" . x, -1)
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

