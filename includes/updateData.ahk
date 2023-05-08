/*	Build call schedule
	Get Docs list
*/

updateCall() {
/*	Update call.xml 
	- Read Qgenda schedule for base Call, Ward, ICU, EP, TXP schedule
	- Read electronic forecast XLS
		\\childrens\files\HCSchedules\Electronic Forecast\2016\11-7 thru 11-13_2016 Electronic Forecast.xlsx
		Move into /root/forecast/call {date=20150301}/<PM_We_F>Del Toro</PM_We_F>
*/
	global y, path, callChg
	
	if FileExist(".lock") {
		return
	}
	FileOpen(".lock", "W")
	
	if fileexist(".\files\call.xml") {
		y := new XML(".\files\call.xml")
	} else {
		y := new XML("<root/>")
		y.addElement("forecast","/root")
		y.save(".\files\call.xml")
	}
	
	callChg := false
	readQgenda()																		; Read Qgenda once daily
	readForecast()																		; Check for Electronic Forecast changes each time
	
	if (callChg=true) {
		progress, , Updating schedules, Syncing...
		dest := "pedcards@homer.u.washington.edu:public_html/patlist/call.xml"
		Run .\files\pscp.exe -sftp -i .\files\trriq-pr.ppk -p .\files\call.xml %dest%,, Min
		sleep 500																		; Citrix VM needs longer delay than 200ms to recognize window
		ConsWin := WinExist("ahk_class ConsoleWindowClass")								; get window ID
		IfWinExist ahk_id %consWin% 
		{
			ControlSend,, {y}{Enter}, ahk_id %consWin%									; blindly send {y}{enter} string to console
		}
		eventlog("Uploaded call list.")
	}
	FileDelete, .lock
	FileCopy, .\files\call.xml, % path.chip "call.xml" , 1
	
	return
}

readForecast() {
	global y, path
	
	; Find the most recently modified "*Electronic Forecast.xls" file
	eventlog("Check electronic forecast.")
	progress,, Updating schedules, Scanning forecast files...
	
	fcLast :=
	fcNext :=
	fcFile := 
	fcFileLong := 
	fcRecent :=
	
	dp:=A_Now
	FormatTime, Wday,%dt%, Wday															; Today's day of the week (Sun=1)
	dp += (2-Wday), days																; Get last Monday's date
	tmp := parsedate(dp)
	fcLast := tmp.mm tmp.dd																; date string "0602" from last week's fc
	
	dt:=A_Now
	dt += (9-Wday), days																; Get next Monday's date
	tmp := parsedate(dt)
	fcNext := tmp.mm tmp.dd																; date string "0609" for next week's fc
	
	Loop, Files, % path.forecast . tmp.yyyy "\*Electronic Forecast*.xls*", F			; Scan through YYYY\Electronic Forecast.xlsx files
	{
		fcFile := A_LoopFileName														; filename, no path
		fcFileLong := A_LoopFileLongPath												; long path
		fcRecent := A_LoopFileTimeModified												; most recent file modified
		if InStr(fcFile,"~") {
			continue																	; skip ~tmp files
		}
		d1 := zDigit(strX(fcFile,"",1,0,"-",1,1)) . zDigit(strX(fcFile,"-",1,1," ",1,1))	; zdigit numerals string from filename "2-19 thru..."
		fcNode := y.selectSingleNode("/root/forecast")									; fcNode = Forecast Node
		
		if (d1=fcNext) {																; this is next week's schedule
			tmp := fcNode.getAttribute("next")											; read the fcNode attr for next week DT-mod (0205-20180202155212)
			if ((strX(tmp,"",1,0,"-",1,1) = fcNext) && (strX(tmp,"-",1,1,"",0) = fcRecent)) { ; this file's M attr matches last adjusted fcNode next attr
				eventlog(fcFile " already done.")
				continue																; if attr date and file unchanged, go to next file
			}
			fcNode.setAttribute("next",fcNext "-" fcRecent)								; otherwise, this is unscanned
			eventlog("fcNext " fcNext "-" fcRecent)
		} else if (d1=fcLast) {															; matches last Monday's schedule
			tmp := fcNode.getAttribute("last")
			if ((strX(tmp,"",1,0,"-",1,1) = fcLast) && (strX(tmp,"-",1,1,"",0) = fcRecent)) { ; this file's M attr matches last week's fcNode last attr
				eventlog(fcFile " already done.")
				continue																; skip to next if attr date and file unchanged
			}
			fcNode.setAttribute("last",fcLast "-" fcRecent)								; otherwise, this is unscanned
			eventlog("fcLast " fcLast "-" fcRecent)										
		} else {																		; does not match either fcNext or fcLast
			continue																	; skip to next file
		}
		
		Progress,, Updating schedules, % fcFile
		FileCopy, %fcFileLong%, fcTemp.xlsx, 1											; create local copy to avoid conflict if open
		eventlog("Parsing " fcFileLong)
		parseForecast(fcRecent)															; parseForecast on this file (unprocessed NEXT or LAST)
	}
	if !FileExist(fcFileLong) {															; no file found
		EventLog("Electronic Forecast.xlsx file not found!")
	}
	
return
}

parseForecast(fcRecent) {
	global y, path, callChg, fcVals
	
	; Initialize some stuff
	if !IsObject(y.selectSingleNode("/root/forecast")) {								; create if for some reason doesn't exist
		y.addElement("forecast","/root")
	} 
	Forecast_svc := []
	Forecast_val := []
	for key,val in fcVals
	{
		tmpVal := strX(val,"",1,0,":",1)
		tmpStr := strX(val,":",1,1,"",0)
		Forecast_svc.Insert(tmpVal)
		Forecast_val.Insert(tmpStr)
	}
	
	fcArr := readXLSX(A_WorkingDir "\fcTemp.xlsx")
	fcDate := []																		; array of dates

	Loop % fcArr.MaxIndex()																; read ROWS
	{
		rowNum := A_Index
		if (rowNum=1) {
			Continue																	; first row is title, skip
		}
		fcRow := fcArr[rowNum]
		rowName := ""																	; ROW name (service name)

		Loop % fcRow.MaxIndex()															; read COLS
		{
			colNum := A_Index
			cel := fcRow[colNum]
			label := false
			if (RegExMatch(cel,"\b(\d{1,2})\D(\d{1,2})(\D\d{2,4})?\b",tmp)) {			; matches date format
				getVals := true
				if !(tmp3) {															; get today's YYYY if not given
					tmp3 := substr(A_now,1,4)
				}
				tmpDt := RegExReplace(tmp3,"\D") zDigit(tmp1) zDigit(tmp2) 				; tmpDt in format YYYYMMDD
				fcDate[colNum] := tmpDt													; fill fcDate[1-7] with date strings
				if !IsObject(y.selectSingleNode("/root/forecast/call[@date='" tmpDt "']")) {
					y.addElement("call","/root/forecast", {date:tmpDt})					; create node if doesn't exist
				}
				continue																; keep getting col dates but don't get values yet
			}
			if !(getVals) {																; don't start parsing until we have passed date row
				continue
			}
			cel := trim(RegExReplace(cel,"\s+"," "))									; remove extraneous whitespace

			if (colNum=1) {																; first column (e.g. A1) is label column
				if (j:=objHasValue(Forecast_val,cel,"RX")) {							; match index value from Forecast_val
					row_name := Forecast_svc[j]											; get abbrev string from index
				} else {
					row_name := RegExReplace(cel,"(\s+)|[\/\*\?]","_")					; no match, create ad hoc and replace space, /, \, *, ? with "_"
				}
				progress,, Scanning forecast, % row_name
				continue																; results in some ROW NAME, now move to the next column
			}
			
			if !(cel~="[a-zA-Z]") {
				cel := ""
			}
			
			fcNode := "/root/forecast/call[@date='" fcDate[colNum] "']"
			if !IsObject(y.selectSingleNode(fcNode "/" row_name)) {						; create node for service person if not present
				y.addElement(row_name,fcNode)
			}
			y.setText(fcNode "/" row_name, cleanString(cel))							; setText changes text value for that node
			
		}
	}

	y.selectSingleNode("/root/forecast").setAttribute("xlsdate",fcRecent)				; change forecast[@xlsdate] to the XLS mod date
	y.selectSingleNode("/root/forecast").setAttribute("mod",A_Now)						; change forecast[@mod] to now

	loop, % (fcN := y.selectNodes("/root/forecast/call")).length						; Remove old call elements
	{
		k:=fcN.item(A_index-1)															; each item[0] on forward
		tmpDt := k.getAttribute("date")													; date attribute
		tmpDt -= A_Now, Days															; diff dates
		if (tmpDt < -21) {																; save call schedule for 3 weeks (for TRRIQ)
			q := y.selectSingleNode("/root/forecast/call[@date='" k.getAttribute("date") "']")
			q.parentNode.removeChild(q)
		}
	}
	y.save(".\files\call.xml")
	Eventlog("Electronic Forecast " fcRecent " updated.")
	callChg := true
	
	Return
}

readXLSX(file) {
; Read XLSX document into array
	oWorkbook := ComObjGet(file)
	colArr := ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"] 	; array of column letters
	arr := {}
	valsEnd := 0																		; flag when reached the last row
	valsEndNum := 3																		; number of blank rows to signify end

	While (valsEnd<valsEndNum)
	{
		rowNum := A_Index
		Progress, % rowNum
		arr[rowNum] := {}																; create array for row
		rowHasVals := False																; check for empty row
		Loop
		{
			colNum := A_Index
			if (colNum>maxCol) {														; push to furthest col with info
				maxCol:=colNum
			}
			cel := oWorkbook.Sheets(1).Range(colArr[colNum] rowNum).value				; Scan Sheet1 A2.. etc

			if (cel!="") {
				rowHasVals := true
				valsEnd:=0
			}
			if ((colNum=maxCol) && (rowHasVals=false)) {								; blank row
				valsEnd++
			}
			if (valsEnd=valsEndNum) {
				arr.Delete(rowNum-valsEndNum+1,rowNum)									; delete end blank rows
				Break
			}
			if ((colNum=maxCol) && (cel="")) {											; at maxCol and empty, break this cols loop
				Break
			}
			arr[rowNum][colNum] := cel
		}
	}

	oExcel := oWorkbook.Application
	oExcel.DisplayAlerts := false
	oExcel.quit

	Return arr
}

readQgenda() {
/*	Fetch upcoming call schedule in Qgenda
	Parse JSON into call elements
	Move into /lists/forecast/call {date=20150301}/<PM_We_F>Del Toro</PM_We_F>
*/
	global y, path, callChg
	
	fcMod := substr(y.selectSingleNode("/root/forecast").getAttribute("mod"),1,8) 
	if (fcMod = substr(A_now,1,8)) {													; Return if already scanned today
		eventlog("Qgenda already done today.")
		return
	}
	
	t0 := t1 := A_now
	t1 += 14, Days
	FormatTime,t0, %t0%, MM/dd/yyyy
	FormatTime,t1, %t1%, MM/dd/yyyy
	IniRead, q_com, .\files\qgenda.ppk, api, com
	IniRead, q_eml, .\files\qgenda.ppk, api, eml
	
	qg_fc := {"CALL":"PM_We_A"
			, "fCall":"PM_We_F"
			, "EP Call":"EP"
			, "ICU":"ICU_A"
			, "TXP Inpt CICU":"Txp_CICU"
			, "TXP Inpt Floor":"Txp_Floor"
			, "IW":"Ward_A"}
	
	progress, , Updating schedules, Auth Qgenda...
	url := "https://api.qgenda.com/v2/login"
	str := httpGetter("POST",url,q_eml
		,"Content-Type=application/x-www-form-urlencoded")
	qAuth := parseJSON(str)[1]															; MsgBox % qAuth[1].access_token
	
	progress, , Updating schedules, Reading Qgenda...
	url := "https://api.qgenda.com/v2/schedule"
		. "?companyKey=" q_com
		. "&startDate=" t0
		. "&endDate=" t1
		. "&$select=Date,TaskName,StaffLName,StaffFName"
		. "&$filter="
		.	"("
		.		"TaskName eq 'CALL'"
		.		" or TaskName eq 'fCall'"
	;	.		" or TaskName eq 'CATH LAB'"
	;	.		" or TaskName eq 'CATH RES'"
		.		" or TaskName eq 'EP Call'"
	;	.		" or TaskName eq 'Fetal Call'"
		.		" or TaskName eq 'ICU'"
	;	.		" or TaskName eq 'TEE/ECHO'"
	;	.		" or TaskName eq 'TEE Call'"
		.		" or TaskName eq 'TXP Inpt CICU'"
		.		" or TaskName eq 'TXP Inpt Floor'"
	;	.		" or TaskName eq 'TXP Res'"
		.		" or TaskName eq 'IW'"
		.	")"
		.	" and IsPublished"
		.	" and not IsStruck"
		. "&$orderby=Date,TaskName"
	str := httpGetter("GET",url,
		,"Authorization= bearer " qAuth.access_token
		,"Content-Type=application/json")
	
	progress, , Updating schedules, Parsing JSON...
	qOut := parseJSON(str)
	
	progress, , Updating schedules, Updating Forecast...
	Loop, % qOut.MaxIndex()
	{
		i := A_Index
		qDate := parseDate(qOut[i,"Date"])												; Date array
		qTask := qg_fc[qOut[i,"TaskName"]]												; Call name
		qNameF := qOut[i,"StaffFName"]
		qNameL := qOut[i,"StaffLName"]
		if (qNameL~="^[A-Z]{2}[a-z]") {													; Remove first initial if present
			qNameL := SubStr(qNameL,2)
		}
		if (qNameL~="Mallenahalli|Chikkabyrappa") {										; Special fix for Sathish and his extra long name
			qNameL:="Mallenahalli Chikkabyrappa"
		}
		if (qnameF qNameL = "JoshFriedland") {											; Special fix for Josh who is registered incorrectly on Qgenda
			qnameL:="Friedland-Little"
		}
		
		if !IsObject(y.selectSingleNode("/root/forecast/call[@date='" qDate.YMD "']")) {
			y.addElement("call","/root/forecast", {date:qDate.YMD})						; create node if doesn't exist
		}
		
		fcNode := "/root/forecast/call[@date='" qDate.YMD "']"
		if !IsObject(y.selectSingleNode(fcNode "/" qTask)) {							; create node for service person if not present
			y.addElement(qTask,fcNode)
		}
		y.setText(fcNode "/" qTask, qNameF " " qNameL)									; setText changes text value for that node
		y.selectSingleNode("/root/forecast").setAttribute("mod",A_Now)					; change forecast[@mod] to now
	}
	
	y.save(".\files\call.xml")
	Eventlog("Qgenda " t0 "-" t1 " updated.")
	callChg := true
	
return
}

getCall(dt) {
	global y
	callObj := {}
	Loop, % (callDate:=y.selectNodes("/root/forecast/call[@date='" dt "']/*")).length {
		k := callDate.item(A_Index-1)
		callEl := k.nodeName
		callVal := k.text
		callObj[callEl] := callVal
	}
	return callObj
}

cleanString(x) {
	replace := {"{":"["															; substitutes for common error-causing chars
				,"}":"]"
				, "\":"/"
				,chr(241):"n"}
				
	for what, with in replace													; convert each WHAT to WITH substitution
	{
		StringReplace, x, x, %what%, %with%, All
	}
	
	x := RegExReplace(x,"[^[:ascii:]]")											; filter remaining unprintable (esc) chars
	
	StringReplace, x,x, `r`n,`n, All										; convert CRLF to just LF
	loop																		; and remove completely null lines
	{
		StringReplace x,x,`n`n,`n, UseErrorLevel
		if ErrorLevel = 0	
			break
	}
	
	return x
}

splitIni(x, ByRef y, ByRef z) {
	y := trim(substr(x,1,(k := instr(x, "="))), " `t=")
	z := trim(substr(x,k), " `t=""")
	return
}

httpGetter(RequestType:="",URL:="",Payload:="",Header*) {
/*	more sophisticated WinHttp submitter, request GET or POST
 *	based on https://autohotkey.com/boards/viewtopic.php?p=135125&sid=ebbd793db3b3d459bfb4c42b4ccd090b#p135125
 */
	hdr := { "form":"application/x-www-form-urlencoded"
			,"json":"application/json"
			,"html":"text/html"}
	
	pWHttp := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	pWHttp.Open(RequestType, URL, 0)
	
	loop, % Header.MaxIndex()
	{
		splitIni(Header[A_index],hdr_type,hdr_val) 
		;~ MsgBox % "'" hdr_type "'`n'" hdr_val "'"
		pWHttp.SetRequestHeader(hdr_type, hdr_val)
	}
	
	if (StrLen(Payload) > 0) {
		pWHttp.Send(Payload)	
	} else {
		pWHttp.Send()
	}
	
	pWHttp.WaitForResponse()
	vText := pWHttp.ResponseText
return vText
}

parseJSON(txt) {
	out := {}
	Loop																		; Go until we say STOP
	{
		ind := A_index															; INDex number for whole array
		ele := strX(txt,"{",n,1, "}",1,1, n)									; Find next ELEment {"label":"value"}
		if (n > strlen(txt)) {
			break																; STOP when we reach the end
		}
		sub := StrSplit(ele,",")												; Array of SUBelements for this ELEment
		Loop, % sub.MaxIndex()
		{
			StringSplit, key, % sub[A_Index] , : , `"							; Split each SUB into label (key1) and value (key2)
			out[ind,key1] := key2												; Add to the array
		}
	}
	return out
}

