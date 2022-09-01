/*	PrevGrab - Grabs data from Preventice website
	Saves results for TRRIQ
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force  ; only allow one running instance per user
#Include %A_ScriptDir%\includes
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

Config: 
{
/*	Set config vals for script
*/
	global webstr:={}
		,  gl:={}
	
	progress,,% " ", Preventice Web

	IfInString, A_ScriptDir, AhkProjects 
	{
		gl.isDevt := true
	} else {
		gl.isDevt := false
	}

	; A_Args[1] := "ftp"				;*******************************

	gl.TRRIQ_path := A_ScriptDir
	gl.files_dir := gl.TRRIQ_path "\files"
	wq := new XML(gl.TRRIQ_path "\worklist.xml")
	
	gl.settings := readIni("settings")
	
	gl.enroll_ct := 0
	gl.inv_ct := 0
	gl.t0 := A_TickCount
}

MainLoop:
{
	eventlog("PREVGRAB: Initializing.")
	progress,,Initializing webdriver...
	
	loop, 3
	{
		eventlog("PREVGRAB: Browser open attempt " A_index)
		wb := wbOpen()																	; start/activate an Chrome/Edge instance
		if IsObject(wb) {
			break
		}
	}
	if !IsObject(wb) {
		eventlog("PREVGRAB: Failed to open browser.")
		progress, hide
		MsgBox, 262160, , Failed to open browser
		ExitApp
	}
	wb.visible := gl.settings.isVisible													; for progress bars
	wb.capabilities.HeadlessMode := gl.settings.isHeadless								; for Chrome/Edge window
	gl.Page := wb.NewSession()															; Session in gl.Page

	if (A_Args[1]="ftp") {
		webStr.FTP := readIni("str_ftp")
		gl.login := readIni("str_ftpLogin")

		PreventiceWebGrab("ftp")
		gl.FAIL := gl.wbFail
	} else {
		; webStr.Enrollment := readIni("str_Enrollment")
		webStr.Inventory := readIni("str_Inventory")
		gl.login := readIni("str_Login")

		; PreventiceWebGrab("Enrollment")
		PreventiceWebGrab("Inventory")
		if (gl.inv_ct < gl.inv_tot) {
			gl.FAIL := true
		}
		filedelete, % gl.files_dir "\prev.txt"											; writeout each one regardless
		FileAppend, % prevtxt, % gl.files_dir "\prev.txt"
		eventlog("PREVGRAB: Enroll " gl.enroll_ct ", Inventory " gl.inv_ct ". (" round((A_TickCount-gl.t0)/1000,2) " sec)")
		
	}
	
	if (gl.FAIL) {																		; Note when a table had failed to load
		MsgBox,262160,, Downloads failed.
		eventlog("PREVGRAB: Critical hit.")
	} else {
		MsgBox,262160,, Preventice update complete!
	}
	
	wbClose()
	gl.Page.Exit()
	wb.driver.Exit()

	ExitApp
}

PreventiceWebGrab(phase) {
	web := webStr[phase]
	
	progress,,% " ",% phase

	wbUrl(web.url)																		; load URL, return DOM in gl.Page
	if (gl.wbFail) {
		return
	}
	wbWaitBusy(gl.settings.webwait)
	prvFunc := web.fx
	loop
	{
		progress,,% "Page " A_index,

		tbl := gl.Page.getElementsByClassName(web.tbl)[0]								; get the Main Table
		if !IsObject(tbl) {
			eventlog("PREVGRAB: *** " phase " *** No matching table.")
			gl.FAIL := true
			return
		}
		
		body := tbl.getElementsByClassName(web.tblBody)[0]
		
		done := %prvFunc%(body)		; parsePreventiceEnrollment() or parsePreventiceInventory() or parsePreventiceFTP()
		
		if (done=0) {																	; no new records returned
			break
		}
		
		PreventiceWebPager(phase,web.changed,web.btn)
	}
	
	gl.Page.Close()																		; release Session object
	return
}

PreventiceWebPager(phase,chgStr,btnStr) {
	pg0 := gl.Page.getElementById(chgStr).innerText
	
	if (tot0 := stRegX(pg0,"i)items.*? of ",1,1,"\d+",0)) {
		gl.inv_tot := tot0
	}

	if (phase="Enrollment") {
		pgNum := gl.enroll_ct
		gl.Page.getElementById(btnStr).click() 											; click when id=btnStr
	}
	if (phase="Inventory") {
		pgNum := gl.inv_ct
		progCt := gl.inv_tot
		if (gl.Page.getElementsByClassName(btnStr)[0].getAttribute("onClick") ~= "return") {
			return
		}
		gl.Page.getElementsByClassName(btnStr)[0].click() 								; click when class=btnstr
	}
	
	t0 := A_TickCount
	While (A_TickCount-t0 < gl.settings.webwait)
	{
		progress,% 100*pgNum/progCt
		
		pg := gl.Page.getElementById(chgStr).innerText
		if (pg != pg0) {
			t1:=A_TickCount-t0
			eventlog("PREVGRAB: " phase " " pgNum " pager (" round(t1/1000,2) " s)"
					, (t1>5000) ? 1 : 0)
			return
		}
		sleep 50
	}
	eventlog("PREVGRAB: " phase " " pgNum " timed out! (" round((A_TickCount-t0)/1000,2) " s)")
	return
}

parsePreventiceFTP(tbl) {
/*	Read FTP table
	Sort by date
	Retrieve the last 1-2 weeks of records
*/
	maxTick := 120000
	dlPath := A_WorkingDir "\pdfTemp"
	gl.Page.CDP.Call("Browser.setDownloadBehavior", { "behavior" : "allow", "downloadPath" : dlPath}) 
	Progress,,% " ",FTP page loaded

	hdr := gl.Page.querySelector("div.table-header-wrapper")
	btn := hdr.querySelectorAll("div[ng-click]")
	loop, % btn.Count()
	{
		k := btn[A_index-1]
		if (k.InnerText ~= "i)Date") {
			btnDate := k
			break
		}
	}

	Progress, 100,% " ",FTP checking sort order
	loop, 2
	{
		gl.Page.tbl := gl.Page.querySelector(".table-body")								; find div with class "table-body"
		gl.Page.tblRows := tbl.querySelectorAll(".row-wrap")							; all rows with class "row-wrap"

		if (ftpDateDiff(0)<7)&&(ftpDateDiff(1)<7) {										; check dates of first 2 rows
			break																		; skip out of they are both within 7 days
		}
		btnDate.click()																	; click to sort list by btnDate
		gl.Page.await()
	}
	sleep 100

	Progress,0,% " ",Parsing FTP list
	ftpList := {}
	loop
	{
		num := A_Index-1
		cols := gl.Page.tblRows[num].querySelectorAll(".ng-binding")
		btnName := cols[0].innertext
		btnDate := cols[2].innertext
		Progress, % 100*A_Index/40, % btnName

		ftpList[num] := btnName
		if (dateDiff(parseDate(btnDate).YMD)) > 21 {
			break
		}
		if (A_index > 100) {
			break
		}
	}
	eventlog("PREVGRAB: Found " ftpList.length() " PDF files.")

	Progress,0,Please be patient...,Fetching PDF files
	loop, read, .\files\mortaras.txt
	{
		k := A_LoopReadLine
		if (k="") {
			Break
		}
		nm:=StrSplit(k,",")
		nm.bestScore := 2
		loop, % ftpList.length()
		{
			rowName := ftpList[A_Index-1]
			rowL := FuzzySearch(Format("{:U}",nm.1),Format("{:U}",rowName))
			rowF := FuzzySearch(Format("{:U}",nm.2),Format("{:U}",rowName))
			
			nm.score := rowL+rowF
			if (nm.score < nm.bestScore) {
				nm.bestScore := nm.score
				nm.bestNum := A_Index-1
				nm.bestName := rowName
			}
		}
		eventlog("PREVGRAB: List name: " k ". "
			. "FTP file: " ftpList[nm.bestNum] " "
			. "(score " round(100*(2-nm.bestScore)/2,2) ")")
		if (nm.bestScore>0.3) {															; skip if match less than 85%
			badFtp .= k "`n"
			continue
		}
		ftpGot := true
		cols := gl.Page.tblRows[nm.bestNum].querySelectorAll(".ng-binding")
		btnName := cols[0]
		btnName.click()
		sleep 200
	}
	if (badFtp) {
		MsgBox 0x10, Missing FTP files
			, % "Could not find PDF files for these patients:`n`n"
			. k "`n`n"
			. "Please check the https://ftp.preventice.com site`n"
			. "and contact Preventice support as needed."
	}
	if !(ftpGot) {
		progress, hide
		return 0
	}

	t0 := A_TickCount
	while (A_TickCount-t0 < 5000) {														; wait for .crdownload to begin
		if FileExist(dlPath "\*crdownload") {
			Break
		}
	}
	eventlog("PREVGRAB: " A_TickCount-t0 " msec to start download.")

	t0 := A_TickCount
	while FileExist(dlPath "\*crdownload") {											; wait for .crdownload to finish
		t1 := A_TickCount-t0
		if (t1 > maxTick) {
			Break
		}
		tbar := SubStr(round(t1/100),-2)
		progress, % tbar
	}
	eventlog("PREVGRAB: " A_TickCount-t0 " msec to download file(s).")

	Progress, Hide
	Return 0
}

checkFtpRow(num=0) {
	col := gl.Page.tblRows[num].querySelectorAll(".ng-binding")							; all div with class "ng-binding"
	dt := col[2].innertext
	nm := col[0].innertext

	return {name:nm,date:dt}
}

ftpDateDiff(row) {
	dt := checkFtpRow(row).date															; get date value from this table row number
	diff := dateDiff(dt)
	return diff
}

parsePreventiceEnrollment(tbl) {
	global prevtxt, gl, wq
	
	lbl_rx := {"demo":"MRN:","dev":"Location:","prov":"GB-SCH-"}						; regex to find blocks
	lbl_pre := {"demo":"NAME:","dev":"SERIAL:","prov":"PROVIDER:"}						; prefix to attach to strings
	
	lbl_demo := {"name":"NAME:","mrn":"MRN:","date":"Created Date:"}					; regex for necessary fields
	lbl_dev := {"dev":"SERIAL:"}
	lbl_prov := {"prov":"PROVIDER:"}
	
	done := 0
	checkdays := gl.settings.checkdays
	
	loop % (trows := tbl.getElementsByTagName("tr")).length								; loop through rows
	{
		r_idx := A_index-1
		trow := trows[r_idx]
		if (trow.getAttribute("id")="") {												; skip the buffer rows
			continue
		}
		res := []
		loop % (tcols := trow.getElementsByTagName("td")).length						; loop through cols
		{
			c_idx := A_Index-1
			txt := tcols[c_idx].innertext
			type := ObjHasValue(lbl_rx,txt,1)											; get type of cell based on regex object
			txt := lbl_pre[type] " " txt "`n"										
			
			for key,val in lbl_%type%													; loop through expected fields in lbl_type
			{
				i := stregX(txt,val,1,1,"\R",1)											; string between lbl and \R
				res[key] := trim(i,": `r`n")
			}
		}
		
		res.name := format("{:U}",parsename(res.name).lastfirst)
		date := parseDate(res.date).YMD
		
		if (dateDiff(date)>checkdays) {													; if days > threshold, break loop
			break
		} else {																		; otherwise done+1 == keep paging
			done ++
		}
		
		prevtxt := "enroll|" 															; prepends enroll item so will be read in chronologic
			. date "|"																	; rather than reverse chronologic order
			. res.name "|"
			. res.mrn "|"
			. res.dev "|"
			. res.prov "|"
			. A_now "`n"
			. prevtxt
		
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
	global prevtxt, gl, wq
	
	gl.clip := tbl.innertext
	if (gl.clip=gl.clip0) {																; no change since last clip
		Return false
	}

	lbl := ["button","model","ser"]
	
	trows := tbl.getElementsByName("tr")
	loop % trows.length()+1																; loop through rows
	{
		r_idx := A_index-1
		trow := trows[r_idx]
		tcols := trow.getElementsByName("td")
		res := []
		loop % lbl.length()																; loop through cols
		{
			c_idx := A_Index-1
			res[lbl[A_index]] := trim(tcols[c_idx].innertext)
		}
		gl.inv_ct ++
		
		if IsObject(wq.selectSingleNode("/root/pending/enroll"
					. "[dev='" res.model " - " res.ser "']")) {							; exists in Pending
			eventlog("PREVGRAB: " res.model " - " res.ser " - already in use.",0)
			continue
		}
		
		prevtxt .= "dev|" res.model "|" res.ser "`n"
	}
	gl.clip0 := gl.clip																	; set the check for repeat copy

	return true
}

wbOpen() {
/*	Use Rufaydium class https://github.com/Xeo786/Rufaydium-Webdriver
	to use Google Chrome or Microsoft Edge webdriver to retrieve webpage
*/
	FileGetVersion, cr32Ver, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
	FileGetVersion, cr64Ver, C:\Program Files\Google\Chrome\Application\chrome.exe
	FileGetVersion, mseVer, C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
	if (cr64Ver) {
		verNum := cr64Ver
		driver := "chromedriver"
		eventlog("PREVGRAB: Found Chrome (x64) version " verNum)
	} Else
	if (cr32Ver) {
		verNum := cr32Ver
		driver := "chromedriver"
		eventlog("PREVGRAB: Found Chrome (x86) version " verNum)
	} Else
	if (mseVer) {
		verNum := mseVer
		driver := "msedgedriver"
		eventlog("PREVGRAB: Found Edge (x86) version " verNum)
	} Else {
		eventlog("PREVGRAB: Could not find installed Chrome or Edge.")
		Return
	}
	Num :=  strX(verNum,"",0,1,".",1,1)

	exe := A_ScriptDir "\files\" driver "\" Num "\" driver ".exe"
	if !FileExist(exe) {
		eventlog("PREVGRAB: Could not find matching driver. Attempt download.")
	}
	wb := new Rufaydium(exe)

	return wb
}

wbUrl(url) {
/*	Open a URL
*/
	gl.wbFail := true
	loop, 3																				; Number of attempts to permit redirects
	{
		try 
		{
			progress,,% "Launching URL attempt " A_Index
			eventlog("PREVGRAB: Navigating to " url " (attempt " A_index ").")
			gl.Page.Navigate(url) 														; load URL
			if !(wbWaitBusy(gl.settings.webwait)) {										; msec before fails
				eventlog("PREVGRAB: Failed to load.")
			}
			if !(gl.Page.URL = url) {
				eventlog("PREVGRAB: Redirected.",0)
				sleep 1000
			}
		}
		catch e
		{
			eventlog("PREVGRAB: wbUrl failed with msg: " stregX(e.message "`n","",1,0,"[\r\n]+",1))
			if instr(e.message,"The RPC server is unavailable") {
				eventlog("PREVGRAB: Reloading DOM...")
				wb := wbOpen()
			}
			continue
		}
		
		if instr(gl.Page.URL,gl.login.string) {
			progress,,% "Sending login"
			loginErr := preventiceLogin()
			eventlog("PREVGRAB: Login " ((loginErr) ? "submitted." : "attempted."))
			sleep 1000
		}
		if (gl.Page.URL=url) {
			eventlog("PREVGRAB: Succeeded.",0)
			gl.wbFail := false
			return
		}
		else {
			eventlog("PREVGRAB: Landed on " gl.Page.URL)
			sleep 500
		}
	}
	gl.wbFail := true
	eventlog("PREVGRAB: Failed all attempts " url)
	return
}

wbClose() {
	gl.Page.exit()
	
	eventlog("PREVGRAB: Closing webdriver.",0)
	
	return
}

wbWaitBusy(maxTick) {
	startTick:=A_TickCount
	
	while InStr(gl.Page.html,"table-body ng-hide") {									; class="table-body ng-hide" present while rendering FTP list
		if (A_TickCount-startTick > maxTick) {
			eventlog("PREVGRAB: " gl.Page.url " timed out.")
			return false
		}
		sleep 500
	} 
	while InStr(gl.Page.html,"{{progress}}") {											; present when rendering FTP login page
		if (A_TickCount-startTick > maxTick) {
			eventlog("PREVGRAB: " gl.Page.url " timed out.")
			return false
		}
		sleep 500
	} 
	while gl.Page.IsLoading() {															; wait until done loading
		if (A_TickCount-startTick > maxTick) {
			eventlog("PREVGRAB: " gl.Page.url " timed out.")
			return false																; break loop if time exceeds maxTick
		}
		checkBtn("Message from webpage","OK")											; check if err window present and click OK button
		sleep 200
	}
	return A_TickCount-startTick
}

checkBtn(txt,btn) {
	if (errHWND:=WinExist(txt)) {
		ControlClick,%btn%,ahk_id %errHWND%
		eventlog("PREVGRAB: Message dialog clicked '" btn "'.")
		sleep 200
	}
	return
}

preventiceLogin() {
/*	Need to populate and submit user login form
*/
	gl.Page
		.getElementById(gl.login.attr_user)
		.value := gl.login.user_name
	
	gl.Page
		.getElementById(gl.login.attr_pass)
		.value := gl.login.user_pass
	
	gl.Page
		.getElementByID(gl.login.attr_btn)
		.click()
	
	if !(wbWaitBusy(gl.login.webwait)) {												; wait until done loading
		return false
	}
	else {
		return true
	}
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
		date.mm := zdigit(d2)
		date.mmm := mo[d2]
		date.dd := zdigit(d3)
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
			, YMDHMS:date.yyyy date.mm date.dd zDigit(time.hr) zDigit(time.min) zDigit(time.sec)
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, MMDD:date.mm "/" date.dd
			, hrmin:zdigit(time.hr) ":" zdigit(time.min)
			, days:zdigit(time.days)
			, hr:zdigit(time.hr), min:zdigit(time.min), sec:zdigit(time.sec)
			, ampm:time.ampm, time:time.time
			, DHM:zdigit(time.days) ":" zdigit(time.hr) ":" zdigit(time.min) " (DD:HH:MM)" 
 			, DT:date.mm "/" date.dd "/" date.yyyy " at " zdigit(time.hr) ":" zdigit(time.min) ":" zdigit(time.sec) }
}

zDigit(x) {
; Add leading zero to a number
	return SubStr("0" . x, -1)
}

dateDiff(d1, d2:="") {
/*	Return date difference in days
	AHK v1L uses envadd (+=) and envsub (-=) to calculate date math
*/
	diff := ParseDate(d2).ymd															; set first date
	diff -= ParseDate(d1).ymd, Days														; d2-d1
	return diff
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

eventlog(event,verbosity:=1) {
/*	verbose 1 or 0 from ini
	verbosity default 1
	verbosity set 0 if only during verbose
*/
	global gl
	
	score := verbosity + gl.settings.verbose
	if (score<1) {
		return
	}
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
	IniRead,x,% gl.files_dir "\prevgrab.ini",%section%
	Loop, parse, x, `n,`r																; analyze section struction
	{
		i := A_LoopField
		if (i~="^(?<!"")\w+:") {														; starts with abc: is an object
			i_type.obj := true
		}
		else if (i~="^(?<!"")\w+=") {													; starts with abc= is a var declaration
			i_type.var := true
		}
		else {																			; anything else is an array list
			i_type.arr := true
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

#Include xml.ahk
#Include sift3.ahk
#Include Rufaydium.ahk