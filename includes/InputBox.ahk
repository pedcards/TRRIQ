InputBox(ByRef OutputVar="", Title="", Text="", Default="`n", Keystrokes="") ;http://www.autohotkey.com/forum/viewtopic.php?p=467756
{
	Static KeysToSend, PID, HWND, PreviousEntries
	If (A_ThisLabel <> "InputBox") {
		If HWND
			SetTimer, InputBox, Off
		If !PID {
			Process, Exist
			PID := ErrorLevel
		}
		If Keystrokes
			KeysToSend := Keystrokes
		WinGet, List, List, ahk_class #32770 ahk_pid %PID%
		HWND = `n0x0`n
		Loop %List%
			HWND .= List%A_Index% "`n"
		If InStr(Default, "`n") and (UsePrev := True)
			StringReplace, Default, Default, `n, , All
		If (Title = "")
			Title := SubStr(A_ScriptName, 1, InStr(A_ScriptName, ".") - 1) ": Input"
		SetTimer, InputBox, 20
		StringReplace, Text, Text, `n, `n, UseErrorLevel
		InputBox, CapturedOutput, %Title%, %Text%, , , Text = "" ? 100 : 116 + ErrorLevel * 18 , , , , , % UsePrev and (t := InStr(PreviousEntries, "`r" (w := (u := InStr(Title, " - ")) ? SubStr(Title, 1, u - 1) : Title) "`n")) ? v := SubStr(PreviousEntries, t += (u ? u - 1 : StrLen(Title)) + 2, InStr(PreviousEntries "`r", "`r", 0, t) - t) : Default
		If !(Result := ErrorLevel) {
			OutputVar := CapturedOutput
			If t
				StringReplace, PreviousEntries, PreviousEntries, `r%w%`n%v%, `r%w%`n%OutputVar%
			Else
				PreviousEntries .= "`r" w "`n" OutputVar
		}
		Return Result
	} Else If InStr(HWND, "`n") {
		If !InStr(HWND, "`n" (TempHWND := WinExist("ahk_class #32770 ahk_pid " PID)) "`n") {
			WinDelay := A_WinDelay
			SetWinDelay, -1
			WinSet, AlwaysOnTop, On, % "ahk_id " (HWND := TempHWND)
			WinActivate, ahk_id %HWND%
			If KeysToSend {
				WinWaitActive, ahk_id %HWND%, , 1
				If !ErrorLevel
					SendInput, %KeysToSend%
				KeysToSend =
			}
			SetTimer, InputBox, -400
			SetWinDelay, %WinDelay%
		}
	} Else If WinExist("ahk_id " HWND) {
		WinSet, AlwaysOnTop, On, ahk_id %HWND%
		SetTimer, InputBox, -400
	} Else
		HWND =
	Return
	InputBox:
	Return InputBox()
}
