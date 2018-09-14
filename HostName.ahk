;****************************************************************************************
; Module: HostName
; Purpose: Determine which clinic an SCH workstation/vdi is installed in.  If the script
;          does not find the workstation/vdi in the xml file (e.g. first time the workstation or
;          vdi has run the script) the script will prompt the user for their location and persist
;          the location to the xml file
;
; Assumptions:
;    -- In general, workstations do not change locations
;    -- Location data for the xml file is up to date (includes all valid locations)
;    -- XML hierarchy: root -> locations
;                      root -> workstations -> workstation -> wsname
;                      root -> workstations -> workstation -> location
;

; **** Globals used for GUI
global SelectedLocation := ""      
global SelectConfirm := ""

; **** Globals used as constants (do not change these variables in the code)
global m_strXmlFilename := "wkslocation.xml"                                 ; path to xml data file that contains workstation information
global m_strXmlLocationsPath := "/root/locations"                            ; xml path to locations node (location names)
global m_strXmlWorkstationsPath := "/root/workstations"                      ; xml path to workstations node (contains all infomation for workstations)
global m_strXmlWksNodeName := "workstation"                                  ; name of "workstation" node in the xml data file
global m_strXmlWksName := "wsname"                                           ; name of the "workstation name node" in the xml data file
global m_strXmlLocationName := "location"                                    ; name of the "location" node in the xml data file

;******************************************************************************
; Fuction: GetLocation
; Purpose: Retrieves the location for the current workstation
; Output : String containing the workstation's location
; Input  : N/A
;
GetLocation()
{
	wks := A_ComputerName
	location := GetWksLocation(wks)
	return location
}

;******************************************************************************
; Function: GetWksLocation
; Purpose : Retrieve the location for the specified workstation from the xml
;           file
; Output  : Success = string containing workstation's name
;           Falure = empty string
; Input   : nameIn - string containing the hostname for the current workstation
; Assumptions : 
;     - xml file named wkslocation.xml
;     - xml file in same folder as script
;     - xml hierarchy is known and static
;     - if workstation is not found user should be prompted for location
;     - return empty string on failure
;
GetWksLocation(nameIn)
{
	location := ""                                                                    ; assume failure

	if FileExist(m_strXmlFilename) {
		locationData := new xml(m_strXmlFilename)                               	  ; load xml file
		wksList := locationData.SelectSingleNode(m_strXmlWorkstationsPath)            ; retreive list of all workstations
		wksFound := false
		loop, % (wsNodes := wksList.SelectNodes(m_strXmlWksNodeName)).Length {        ; loop through the workstations
			wsInfoNode := wsNodes.item(A_Index - 1)                                   ; Retrieve workstation node from workstation list
			wsName     := wsInfoNode.SelectNodes(m_strXmlWksName).item(0).Text        ; Retrieve the wsname node from the workstation informaton
				
			if (wsName = nameIn)                                                      ; compare this workstation to current workstation list name
			{
				location := wsInfoNode.SelectNodes(m_strXmlLocationName).item(0).Text ; names matched, retreive the workstation location
				wksfound := true
				break
			}
		}
	
		if (wksfound = false) {
			location := PromptForLocation()                                           ; Prompt user for location of new workstation
		}
	
	} else {
			MsgBox, 16, File Error, Location data unavailable: The file %m_strXmlFilename% was not found.
	}
	return location
}
	
;******************************************************************************
; Function: PromptForLocation
; Purpose : Retrieve the location of the current workstation from the user by
;           displaying a dialog that will allow the user to select their loctaion 
;           from a list of all the available locations.  After user selects the location
;           call the function to persist the location to the data store.
;
PromptForLocation()
{
	workstationLocation := ""

    locationData := GetLocations()                                              ; Function to retrive the location list from the data store

	;Buld and display the dialog box
	Gui, New, AlwaysOnTop -MaximizeBox -MinimizeBox, Unknown Location
	Gui, Add, Text, x15 y20 w250 h60,The application is unable to determine your location. Please select your location from the list and confirm that you made the correct selection.
	Gui, Add, ListBox, vSelectedLocation x15 y70 w245 h200 Sort gLocationList_Click, %locationData%
	Gui, Add, Text,x100 y290 w72, You selected:
	Gui, Add, Text, vSelectConfirm x172 y290 w150, %SelectConfirm% 
	Gui, Add, Button, x160 y315 w100 gConfirmBtn_Click, Confirm
    Gui, Show, w275 h350
	
	WinWaitClose, Unknown Location                                              ;wait for the user to respond
	return %workstationLocation%                                                ;return the selected location

	;******************* Gui Event handlers (subroutines) *********************
	LocationList_Click:
		Gui, Submit, Nohide                                                     ; user selected location from list, submit dialog data / keep displaying the dialog
		GuiControl,, SelectConfirm, %SelectedLocation%                          ; reflect selected value in confirmation text box
	return

	ConfirmBtn_Click:
		Gui, Submit, Hide                                                       ; User made/confirmed selection, submit data
		AddWorkstation(SelectedLocation)                                        ; Persist workstation/location to data store
		workstationLocation := SelectedLocation                                 ; set the return value to the selected location
		WinClose, Unknown Location                                              ; Close the dialog
		Gui, Destroy                                                            ; Release resources
	return
}

;******************************************************************************
; Function: GetLocations
; Purpose : Retrieve the location list from the data store in a format compatible
;           for use in a Gui ListBox control
; Output  : String contining piped list of locations
; Input   : N/A
;
GetLocations()
{
	locationList = ""
	
	locationData := new xml(m_strXmlFilename)                       ; Read xml file
	
	wksList := locationData.SelectSingleNode(m_strXmlLocationsPath)      ; Retreive Locations node
	loop, % (wksNodes := wksList.SelectNodes(m_strXmlLocationName)).Length     ; Loop through node and create piped list of locations
	{
		location:= wksNodes.item(A_Index - 1).selectSingleNode("site").text
		if (A_Index = 1) {
			locationList := location                                 ; No pipe symbol before fist location
		} else {
			locationList := locationList . "|" . location
		}
	}
	return %locationList%
}

;******************************************************************************
; Function: AddWorkstation
; Purpose : Persist the workstation/location to the data store
; Output  : N/A
; Input   : locationData - pointer to the 
;
AddWorkstation(location)
{
	locationData := new xml(m_strXmlFilename) 
	
	workstations := locationData.SelectSingleNode(m_strXmlWorkstationsPath)
	workstation := locationData.addChild(m_strXmlWorkstationsPath,"element",m_strXmlWksNodeName)
	
	wsnameNode := locationData.createNode(1,m_strXmlWksName,"")
	wsnameNode.Text := A_ComputerName
	workstation.appendChild(wsnameNode)
	
	locationNode := locationData.createNode(1,m_strXmlLocationName,"")
	locationNode.Text := location
	workstation.appendChild(locationNode)
	
    locationData.TransformXML()
	locationData.saveXML()
}



