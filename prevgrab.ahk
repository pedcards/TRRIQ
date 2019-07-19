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
		, fx:"ParsePreventiceEnrollment"}
	webStr.Inventory := {dlg:"Facility`nInventory Status`nDevice in Hand (Enrollment not linked)"
		, url:"https://secure.preventice.com/Facilities/"
		, win:"Facilities"
		, tbl:"ctl00_mainContent_InventoryStatus_userControl_gvInventoryStatus_ctl00"
		, changed:"ctl00_mainContent_InventoryStatus_userControl_gvInventoryStatus_ctl00_Pager"
		, btn:"rgPageNext"
		, fx:"ParsePreventiceInventory"}
	
	IfInString, A_ScriptDir, AhkProjects 
	{
		gl.isAdmin := true
		gl.TRRIQ_path := A_ScriptDir
	} else {
		gl.isAdmin := false
		gl.TRRIQ_path := "\\childrens\files\HCCardiologyFiles\EP\HoltER Database\TRRIQ"
	}
	gl.files_dir := gl.TRRIQ_path "\files"
	
	loop, read, % gl.files_dir "\prev.key"
	{
		k := A_LoopReadLine
		fld := strX(k,"",0,0,"=",1,1)
		val := strX(k,"=",1,1,"",0)
		gl[fld] := val
	}
	;~ eventlog("PREVGRAB: user (" strlen(gl.user_name) "), pass (" strlen(gl.user_pass) ")")
	
	gl.enroll_ct := 0
	gl.inv_ct := 0
}

MainLoop:
{
	wb := IEopen()																		; start/activate an IE instance
	wb.visible := false
	
	PreventiceWebGrab("Enrollment")
	
	PreventiceWebGrab("Inventory")
	
	filedelete, % gl.files_dir "\prev.txt"
	FileAppend, % prevtxt, % gl.files_dir "\prev.txt"
	
	eventlog("PREVGRAB: Enroll " gl.enroll_ct ", Inventory " gl.inv_ct ".")
	
	IEclose()
	
	ExitApp
}

PreventiceWebGrab(phase) {
	global webStr, wb
	web := webStr[phase]
	
	if (gl.isAdmin) {
		progress,,% " ",% phase
	}
	IEurl(web.url)																		; load URL, return DOM in wb
	prvFunc := web.fx
	
	loop
	{
		tbl := wb.document.getElementById(web.tbl)										; get the Main Table
		if !IsObject(tbl) {
			eventlog("PREVGRAB ERR: *** No matching table.")
			return
		}
		
		body := tbl.getElementsByTagName("tbody")[0]
		clip := body.innertext
		if (clip=clip0) {																; no change since last clip
			break
		}
		
		done := %prvFunc%(body)		; parsePreventiceEnrollment() or parsePreventiceInventory()
		
		if (done=0) {																	; no new records returned
			break
		}
		clip0 := clip																	; set the check for repeat copy
		
		PreventiceWebPager(phase,web.changed,web.btn)
	}
	
	wb.navigate(web.url)																; refresh first page
	ComObjConnect(wb)																	; release wb object
	return
}

PreventiceWebPager(phase,chgStr,btnStr) {
	global wb
	
	if (phase="Enrollment") {
		wb.document.getElementById(btnStr).click() 										; click when id=btnStr
	}
	if (phase="Inventory") {
		wb.document.getElementsByClassName(btnStr)[0].click() 							; click when class=btnstr
	}
	pg0 := wb.document.getElementById(chgStr).innerText
	
	;~ While ((wb.busy) || (wb.ReadyState < 3))
	loop, 200																			; wait up to 100*0.05 = 5 sec
	{
		pg := wb.document.getElementById(chgStr).innerText
		if (gl.isAdmin) {
			progress,,% wb.ReadyState, % phase " (" A_index ")"
		}
		if (pg != pg0) {
			;~ MsgBox Finished!
			break
		}
		sleep 50
	}
	;~ MsgBox Timed out!
	return
}

parsePreventiceEnrollment(tbl) {
	global prevtxt, gl
	
	lbl := ["name","mrn","date","dev","prov"]
	done := 0
	checkdays := 21
	
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
		
		prevtxt .= "enroll|" 
			. date "|"
			. res.name "|"
			. res.mrn "|"
			. res.dev "|"
			. res.prov "|"
			. A_now "`n"
		
		gl.enroll_ct ++
	}
	
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
	global prevtxt, gl
	
	lbl := ["button","model","ser"]
	
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
		prevtxt .= "dev|" res.model "|" res.ser "`n"
		gl.inv_ct ++
	}

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
	while ((wb.busy) ||	(wb.ReadyState < 3)) {											; wait until done loading
		if (gl.isAdmin) {
			progress,,% wb.ReadyState
		}
		sleep 10
	}
	
	if instr(wb.LocationURL,"UserLogin") {
		preventiceLogin()
	}
	
	return
}

IEclose() {
	DetectHiddenWindows, On
	while WinExist("ahk_class IEFrame")
	{
		i := A_index
		Process, Close, iexplore.exe
	}
	eventlog("PREVGRAB: Closed " i " IE windows.")
	
	return
}

preventiceLogin() {
/*	Need to populate and submit user login form
*/
	global wb, gl
	attr_user := "ctl00$PublicContent$Login1$UserName"
	attr_pass := "ctl00$PublicContent$Login1$Password"
	attr_btn := "ctl00$PublicContent$Login1$goButton"	
	
	wb.document
		.getElementById(RegExReplace(attr_user,"\$","_"))
		.value := gl.user_name
	wb.document
		.getElementById(RegExReplace(attr_pass,"\$","_"))
		.value := gl.user_pass
	wb.document
		.getElementByID(RegExReplace(attr_btn,"\$","_"))
		.click()
	while wb.busy {																		; wait until done loading
		sleep 10
	}

	return
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

eventlog(event) {
	global gl
	user := A_UserName
	comp := A_ComputerName
	FormatTime, sessdate, A_Now, yyyy.MM
	FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
	name := gl.TRRIQ_path "\logs\" . sessdate . ".log"
	txt := now " [" user "/" comp "] " event "`n"
	filePrepend(txt,name)
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

