#include <Array.au3>
#include <Excel.au3>
#include <Date.au3>
#include <File.au3>
#include <TrayConstants.au3>

;Second exe file to intergrate into main exe. It Navigates the WAF macro's when they are started.
FileInstall("V:\TSC\K&S\WG Portaal\Rick\GO_eWG_Testscript\AutoIT\QuRickstart\QuRickstart_v10.0\QuRickstart_Temp.exe", @ScriptDir & "/QuRickstart_Temp.exe", 1)

;----------------------& .Ini global variables &----------------------;
;Initialise global variables with the values of the config.ini.
Global $g_sIniFilePath 			= @ScriptDir & "\QuRickstartConfig.ini"
Global $g_sHotKeyStart 			= IniRead($g_sIniFilePath, "Hotkeys"	, "HotKey_Start_Normal"			, "F4")
Global $g_sHotKeyDebug 			= IniRead($g_sIniFilePath, "Hotkeys"	, "HotKey_Start_Debug"			, "F3")
Global $g_sHotKeyStop 			= IniRead($g_sIniFilePath, "Hotkeys"	, "HotKey_Stop_QuRickstart"		, "Escape")
Global $g_sWAFWorkbookPath		= IniRead($g_sIniFilePath, "WAF"		, "WAF_Workbook_Path"			, "")
If $g_sWAFWorkbookPath == "" Then
   MsgBox(0, "QuRickstart", "The path to the WAF workbook is empty, please add the path to the Quickstart.ini file." & @CRLF & @CRLF & "Stopping script..")
   Exit
EndIf
Global $g_bDeleteTestcase 		= IniRead($g_sIniFilePath, "Testcase"	, "Delete_XML_Testcase"			, True)
Global $g_sCellTestscenario 	= IniRead($g_sIniFilePath, "CoverSheet"	, "Excel_Cell_Testscenario"		, "B4")
Global $g_sCellDefaultFolder 	= IniRead($g_sIniFilePath, "CoverSheet"	, "Excel_Cell_DefaultFolder"	, "B9")

;----------------------& Other global variables &----------------------;
;Initialise the other global variables.
Global $g_bDebugModes 			= False
Global $g_oWorkbook 			= Null
Global $g_sTitle 				= ""
Global $g_sDefaultFolder 		= ""
Global $g_sTestscenarioName 	= ""

;----------------------------& Script &-------------------------------;
;Set hotkeys and wait for input.
HotKeySet("{" & $g_sHotKeyStart	& "}", "_StartTestNormal")
HotKeySet("{" & $g_sHotKeyDebug	& "}", "_DebugModes")
HotKeySet("{" & $g_sHotKeyStop 	& "}", "_ExitScript")
_Wait()

;Parameter:		Non.
;Returns: 		Nothing.
;Description: 	Main script.
Func _Main()
   If StringRegExp(WinGetTitle("[ACTIVE]"), "(?i).*Excel.*") == 0 Then ;Check if the opend program is Excel based on the window's title.
	  MsgBox(0, "QuRickstart error!", "Please start the script in an active testscript excelsheet.")
	  _Wait()
   EndIf
   If $g_sTitle <> WinGetTitle("[ACTIVE]") Then ;First time the script is started in this Excelsheet? This If statement gets the Excel workbook object, otherwise connect.
	  $g_sTitle = WinGetTitle("[ACTIVE]")
	  local $sFilename = StringRight($g_sTitle,StringLen($g_sTitle)-18) ;Remove "Microsoft Excel - " form the title.
	  Local $asWorkbookList = _Excel_BookList()
	  If @error Then
		 MsgBox(0, "QuRickstart error!", "Error listing Excel workbooks." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		 _Wait()
	  EndIf
	  For $i = 0 To UBound($asWorkbookList) -1 Step +1
		 If StringRegExp($asWorkbookList[$i][1], $sFilename & ".*") == 1 Then ;Loop through all active workbooks and find the active Excelsheet.
			Local $sFilePath = $asWorkbookList[$i][2] & "\" & $asWorkbookList[$i][1]
			$g_oWorkbook = _Excel_BookAttach($sFilePath)
			If @error Then
			   MsgBox(0, "QuRickstart error!", "Error attaching to workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
			   _Wait()
			EndIf
			$g_sDefaultFolder = _ReadWorkbook($g_oWorkbook, 1, $g_sCellDefaultFolder)
			$g_sTestscenarioName = _ReadWorkbook($g_oWorkbook, 1, $g_sCellTestscenario)
			ExitLoop
		 EndIf
	  Next
   EndIf
   Local $sCurrentSheet = $g_oWorkbook.ActiveSheet.Name
   If StringRegExp($sCurrentSheet, "(?i).*_testdata") == 1 Then
	  MsgBox(0, "QuRickstart error!", "This is not a testcase sheet, please start the script in a testcase.")
   EndIf
   Local $sTestcaseID = _CreateID()
   _CreateXML(_TestcasePoistion($g_oWorkbook, $sCurrentSheet), $sTestcaseID)
   Local $oExcel = _Excel_Open()
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error opening excel" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  _Wait()
   EndIf
   Local $oWorkbookWAF = _Excel_BookOpen($oExcel, $g_sWAFWorkbookPath)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error opening the workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  _Wait()
   EndIf
   Local $iAmounOfPgdn = _CountPgDn($g_sDefaultFolder)
   Run("QuRickstart_Temp.exe " & $iAmounOfPgdn, "") ;Run second .exe file to navigate the WAF excelhseet. It waits for the WAF macro's to start.
   If $g_bDebugModes == True Then ;Start the "StartTestDebug" or "StartTest" macro's of the WAF excelsheet.
	  $oExcel.Run("StartTestDebug")
   Else
	  $oExcel.Run("StartTest")
   EndIf
   If $g_bDeleteTestcase == True Then ;Delete the generated testcase.
	  Local $sTestscasePath = $g_sDefaultFolder & $sTestcaseID & " " & $g_sTestscenarioName & ".xml"
	  Local $iDelete = FileDelete($sTestscasePath)
	  If $iDelete == 0 Then
		 MsgBox(0, "QuRickstart error!", "Testcase '" & $sTestcaseID & " " & $g_sTestscenarioName & "' could not be deleted, I don't know why...")
	  EndIf
   EndIf
   _Wait()
EndFunc

;---------------------------& Functions &---------------------------;
;Parameter: 	Non.
;Returns: 		An ID (String).
;Description: 	Generates an unique ID based on the date and time.
Func _CreateID()
   Local $aiTijd = StringSplit(_NowTime(), ":")
   Local $aiDatum = StringSplit(_NowDate(), "-")
   Local $sUniqueID = "zz" & StringRight($aiDatum[3],2) & $aiDatum[2] & $aiDatum[1] & $aiTijd[1] & $aiTijd[2] & $aiTijd[3]
   Return($sUniqueID)
EndFunc

;Parameter:		1. Number of the testcase.
;				2. Unique ID of the testcase.
;Returns: 		Nothing.
;Description: 	Navigates the testscript excelsheet to validate the testcase and generate the XML file.
Func _CreateXML($sTestcaseNumber, $sTestcaseID)
   Send("!y") ;Validate the XML.
   Send("{y Down}{5}")
   Sleep(100)
   local $sCurrentWindow = WinGetTitle("[ACTIVE]")
   If $sCurrentWindow == "Error Log" Then ;Error in validation? Script on hold.
	  _Wait()
   EndIf
   Send("{ENTER}")
   Send("!y") ;Generate the XML file.
   Send("{y Down}{8}{TAB 5}")
   Local $iNumberInList = -1 + $sTestcaseNumber
   Send("{DOWN " & $iNumberInList & "}")
   Send("{SPACE}{TAB 6}{SPACE}{TAB}{ENTER}{TAB 2}{LEFT}")
   Send($sTestcaseID & " ") ;Types unique ID before the testscenario name.
   Send("{TAB}{ENTER 2}")
EndFunc

;Parameter: 	1. Excel workbook object.
;				2. Excel sheet number.
;				3. Excel cell number.
;Returns: 		The value of a Excel cell (String).
;Description: 	Reads an Excel workbook -> Sheet -> Cell to return the value.
Func _ReadWorkbook($oWorkbook, $sSheet, $sExcelCell)
   Local $sCellValue = _Excel_RangeRead($oWorkbook, $sSheet, $sExcelCell)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  _Wait()
   EndIf
   Return($sCellValue)
EndFunc

;Parameter:		A directory path.
;Returns: 		Number (Integer).
;Description: 	Counts the number of files in a directory and calculates how many times to press the key PgDn to reach the last record in the WAF selection list.
Func _CountPgDn($sDir)
   Local $aFileList = _FileListToArray($sDir)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error listing files in directory: " & $sDir & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  _Wait()
   EndIf
   Local $iFileCount = $aFileList[0]
   If $iFileCount <= 19 Then
	  $iAmounOfPgdn = 1
   Else
	  $iAmounOfPgdn = (Int(($iFileCount - 19) / 18) + 2)
   EndIf
   Return $iAmounOfPgdn
EndFunc

;Parameter:		Excel workbook object.
;Returns:		Number (Integer).
;Description:	Returns the position of the selected testcase relative to other testcases in the testscript.
Func _TestcasePoistion($oWorkbook, $sCurrentSheet)
   Local $asSheetList = _Excel_SheetList($oWorkbook)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error listing worksheets." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  _Wait()
   EndIf
   Local $asSheetListFiltered[1] = ["Init"]
   _ArrayDelete($asSheetListFiltered, 0) ;Delete the initialized value in the array.
   For $i = 0 To UBound($asSheetList) -1 Step +1
	  If StringRegExp($asSheetList[$i][0], "(?i).*_testdata") Or $asSheetList[$i][0] == $sCurrentSheet Then ;Filter out all sheets that contain "_Testdata", it emplies it's a testcase.
		 _ArrayAdd($asSheetListFiltered, $asSheetList[$i][0])
	  EndIf
   Next
   Local $asToDelete = [2, 0, 1]
   _ArrayDelete($asSheetListFiltered, $asToDelete) ;Remove the first 2 standard sheets named "Testcase_Testdata" and "Centrale_Testdata" from array.
   local $iIndexCurrentSheet = _ArraySearch($asSheetListFiltered, $sCurrentSheet)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "This is not a testcase sheet, please start the script in a testcase.")
	  _Wait()
   EndIf
   Return $iIndexCurrentSheet +1 ;Current position of the sheet is the index in the array + 1, because the array starts at 0.
EndFunc

;Parameter:		Non.
;Returns: 		Nothing.
;Description:	Sets the variable $g_bDebugModes to False and starts the main script.
Func _StartTestNormal()
   $g_bDebugModes = False
   TrayTip("QuRickstart", "Starting testcase!",10,1)
   _Main()
EndFunc

;Parameter:		Non.
;Returns: 		Nothing.
;Description: 	Sets the variable $g_bDebugModes to True and starts the main script.
Func _DebugModes()
   $g_bDebugModes = True
   TrayTip("QuRickstart", "Starting testcase in debug mode!",10,1)
   _Main()
EndFunc

;Parameter:		Non.
;Returns: 		Nothing.
;Description: 	Just sleeps forever until a hotkey is pressed.
Func _Wait()
   While True
	  Sleep(100)
   WEnd
EndFunc

;Parameter:		Non.
;Returns: 		Nothing.
;Description: 	Exits the script.
Func _ExitScript()
   TrayTip("QuRickstart", "Stopping QuRickstart.... >:(",10,1)
   Sleep(1000)
   Exit
EndFunc

;Author: Rick Ensink. Rick.Ensink@(uwv.nl/sogeti.com), 0645037152.
;K&S, Team Coconut, Groot Onderhoud E-Werkgever/Zakelijk.
;Date: 06-03-2019.