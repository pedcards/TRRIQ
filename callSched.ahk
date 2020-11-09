/*	Build call schedule
	Extracted from chipotle.ahk
*/

readForecast() {
/*	Read electronic forecast XLS
	\\childrens\files\HCSchedules\Electronic Forecast\2016\11-7 thru 11-13_2016 Electronic Forecast.xlsx
	Move into /lists/forecast/call {date=20150301}/<PM_We_F>Del Toro</PM_We_F>
*/
	global y, path
	
	; Get Qgenda items
	fcMod := substr(y.selectSingleNode("/root/lists/forecast").getAttribute("mod"),1,8) 
	if !(fcMod = substr(A_now,1,8)) {													; Forecast has not been scanned today
		readQgenda()																	; Read Qgenda once daily
	}
	
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
		fcNode := y.selectSingleNode("/root/lists/forecast")							; fcNode = Forecast Node
		
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
	
	Progress, off	
	
return
}

parseForecast(fcRecent) {
	global y, path
		, forecast_val, forecast_svc
	
	; Initialize some stuff
	if !IsObject(y.selectSingleNode("/root/lists/forecast")) {							; create if for some reason doesn't exist
		y.addElement("forecast","/root/lists")
	} 
	colArr := ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"] 	; array of column letters
	fcDate:=[]																			; array of dates
	oWorkbook := ComObjGet(A_WorkingDir "\fcTemp.xlsx")
	getVals := false																	; flag when have hit the Date vals row
	valsEnd := false																	; flag when reached the last row
	
	; Scan through XLSX document
	While !(valsEnd)																	; ROWS
	{
		RowNum := A_Index
		row_nm :=																		; ROW name (service name)
		if (rowNum=1) {																	; first row is title, skip
			continue
		}
		
		Loop																			; COLUMNS
		{
			colNum := A_Index															; next column
			if (colNum=1) {
				label:=true																; first column (e.g. A1) is label column
			} else {
				label:=false
			}
			if (ColNum>maxCol) {														; increment maxCol
				maxCol:=colNum
			}
			
			cel := oWorkbook.Sheets(1).Range(colArr[ColNum] RowNum).value				; Scan Sheet1 A2.. etc
			if ((cel="") && (colnum=maxcol)) {											; at maxCol and empty, break this cols loop
				break
			}
			if (RegExMatch(cel,"\b(\d{1,2})\D(\d{1,2})(\D\d{2,4})?\b",tmp)) {			; matches date format
				getVals := true
				if !(tmp3) {															; get today's YYYY if not given
					tmp3 := substr(A_now,1,4)
				}
				tmpDt := RegExReplace(tmp3,"\D") zDigit(tmp1) zDigit(tmp2) 				; tmpDt in format YYYYMMDD
				fcDate[colNum] := tmpDt													; fill fcDate[1-7] with date strings
				if !IsObject(y.selectSingleNode("/root/lists/forecast/call[@date='" tmpDt "']")) {
					y.addElement("call","/root/lists/forecast", {date:tmpDt})			; create node if doesn't exist
				}
				continue																; keep getting col dates but don't get values yet
			}
			
			if !(getVals) {																; don't start parsing until we have passed date row
				continue
			}
			
			cel := trim(RegExReplace(cel,"\s+"," "))									; remove extraneous whitespace
			if (label) {
				if !(cel) {																; blank label means we've reached the end of rows
					valsEnd := true														; flag to end
					break																; break out of LOOP to next WHILE
				}
				
				if (j:=objHasValue(Forecast_val,cel,"RX")) {							; match index value from Forecast_val
					row_nm := Forecast_svc[j]											; get abbrev string from index
				} else {
					row_nm := RegExReplace(cel,"(\s+)|[\/\*\?]","_")					; no match, create ad hoc and replace space, /, \, *, ? with "_"
				}
				progress,, Scanning forecast, % row_nm
				continue																; results in some ROW NAME, now move to the next column
			}
			if !(cel~="[a-zA-Z]") {
				cel := ""
			}
			
			fcNode := "/root/lists/forecast/call[@date='" fcDate[colNum] "']"
			if !IsObject(y.selectSingleNode(fcNode "/" row_nm)) {						; create node for service person if not present
				y.addElement(row_nm,fcNode)
			}
			y.setText(fcNode "/" row_nm, cleanString(cel))								; setText changes text value for that node
		}
	}
	
	oExcel := oWorkbook.Application
	oExcel.DisplayAlerts := false
	oExcel.quit
	
	y.selectSingleNode("/root/lists/forecast").setAttribute("xlsdate",fcRecent)			; change forecast[@xlsdate] to the XLS mod date
	y.selectSingleNode("/root/lists/forecast").setAttribute("mod",A_Now)				; change forecast[@mod] to now

	loop, % (fcN := y.selectNodes("/root/lists/forecast/call")).length					; Remove old call elements
	{
		k:=fcN.item(A_index-1)															; each item[0] on forward
		tmpDt := k.getAttribute("date")													; date attribute
		tmpDt -= A_Now, Days															; diff dates
		if (tmpDt < -21) {																; save call schedule for 3 weeks (for TRRIQ)
			q := y.selectSingleNode("/root/lists/forecast/call[@date='" k.getAttribute("date") "']")
			q.parentNode.removeChild(q)
		}
	}
	y.save(path.chip "currlist.xml")
	Eventlog("Electronic Forecast " fcRecent " updated.")
Return
}

readQgenda() {
/*	Fetch upcoming call schedule in Qgenda
	Parse JSON into call elements
	Move into /lists/forecast/call {date=20150301}/<PM_We_F>Del Toro</PM_We_F>
*/
	global y, path
	
	t0 := t1 := A_now
	t1 += 14, Days
	FormatTime,t0, %t0%, MM/dd/yyyy
	FormatTime,t1, %t1%, MM/dd/yyyy
	IniRead, q_com, % path.chip "qgenda.ppk", api, com
	IniRead, q_eml, % path.chip "qgenda.ppk", api, eml
	
	qg_fc := {"CALL":"PM_We_A"
			, "fCall":"PM_We_F"
			, "EP Call":"EP"
			, "ICU":"ICU_A"
			, "TXP Inpt":"Txp"
			, "IW":"Ward_A"}
	
	progress, , Updating schedules, Auth Qgenda...
	url := "https://api.qgenda.com/v2/login"
	str := httpGetter("POST",url,q_eml
		,"Content-Type=application/x-www-form-urlencoded")
	qAuth := parseJSON(str)[1]																; MsgBox % qAuth[1].access_token
	
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
		.		" or TaskName eq 'TXP Inpt'"
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
		qDate := parseDate(qOut[i,"Date"])										; Date array
		qTask := qg_fc[qOut[i,"TaskName"]]										; Call name
		qNameF := qOut[i,"StaffFName"]
		qNameL := qOut[i,"StaffLName"]
		if (qNameL~="^[A-Z]{2}[a-z]") {											; Remove first initial if present
			qNameL := SubStr(qNameL,2)
		}
		if (qNameL~="Mallenahalli|Chikkabyrappa") {								; Special fix for Sathish and his extra long name
			qNameL:="Mallenahalli Chikkabyrappa"
		}
		if (qNameL qNameF = "NelsonJames") {									; Special fix to make Tony findable on paging call site
			qNameF:="Tony"
		}
		if (qnameF qNameL = "JoshFriedland") {									; Special fix for Josh who is registered incorrectly on Qgenda
			qnameL:="Friedland-Little"
		}
		
		if !IsObject(y.selectSingleNode("/root/lists/forecast/call[@date='" qDate.YMD "']")) {
			y.addElement("call","/root/lists/forecast", {date:qDate.YMD})		; create node if doesn't exist
		}
		
		fcNode := "/root/lists/forecast/call[@date='" qDate.YMD "']"
		if !IsObject(y.selectSingleNode(fcNode "/" qTask)) {					; create node for service person if not present
			y.addElement(qTask,fcNode)
		}
		y.setText(fcNode "/" qTask, qNameF " " qNameL)							; setText changes text value for that node
		y.selectSingleNode("/root/lists/forecast").setAttribute("mod",A_Now)	; change forecast[@mod] to now
	}
	
	y.save(path.chip "currlist.xml")
	Eventlog("Qgenda " t0 "-" t1 " updated.")
	
return
}

getCall(dt) {
	global y
	callObj := {}
	Loop, % (callDate:=y.selectNodes("/root/lists/forecast/call[@date='" dt "']/*")).length {
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

