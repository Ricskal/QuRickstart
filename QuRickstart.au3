#include <Array.au3>
#include <Excel.au3>
#include <Date.au3>
#include <File.au3>
#include <TrayConstants.au3>

;Second .exe file to navigate the WAF Excelsheet when the test macro's are started.
FileInstall("V:\TSC\K&S\WG Portaal\Rick\GO_eWG_Testscript\AutoIT\QuRickstart\QuRickstart_v9.0\QuRickstart_Temp.exe", @ScriptDir & "/QuRickstart_Temp.exe", 1)

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
Global $g_sCellTestcase 		= IniRead($g_sIniFilePath, "Testcase"	, "Excel_Cell_Testcase"			, "B1")
Global $g_sCellTestscenario 	= IniRead($g_sIniFilePath, "CoverSheet"	, "Excel_Cell_Testscenario"		, "B4")
Global $g_sCellDefaultFolder 	= IniRead($g_sIniFilePath, "CoverSheet"	, "Excel_Cell_DefaultFolder"	, "B9")

;----------------------& Other global variables &----------------------;
;Initialise the other global variables.
Global $g_bDebugModes 			= False
Global $g_oWorkbook 			= Null
Global $g_sTitle 				= ""
Global $g_sDefaultFolder 		= ""
Global $g_sTestscenarioName 	= ""

;------------------------& Script &---------------------------;
;Set hotkeys and wait for input.
HotKeySet("{" & $g_sHotKeyStart	& "}", "_StartTestNormal")
HotKeySet("{" & $g_sHotKeyDebug	& "}", "_DebugModes")
HotKeySet("{" & $g_sHotKeyStop 	& "}", "_ExitScript")
_Wait()

;Input: 		Nothing.
;Returns: 		Nothing.
;Description: 	Main script.
Func _Main()
   If $g_sTitle == WinGetTitle("[ACTIVE]") And StringRegExp($g_sTitle, "(?i).*Excel.*") == 1 Then
   Else
	  If StringRegExp(WinGetTitle("[ACTIVE]"), "(?i).*Excel.*") == 1 Then
		 $g_sTitle = WinGetTitle("[ACTIVE]")
		 local $sFilename = StringRight($g_sTitle,StringLen($g_sTitle)-18)
		 Local $asWorkbookList = _Excel_BookList()
		  If @error Then
			MsgBox(0, "QuRickstart error!", "Error listing Excel workbooks." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
			_Wait()
		 EndIf
		 For $i = 0 To UBound($asWorkbookList) -1
			If StringRegExp($asWorkbookList[$i][1], $sFilename & ".*") == 1 Then
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
	  Else
		 MsgBox(0, "QuRickstart error!", "Please start the script in an active testscript excelsheet.")
		 _Wait()
	  EndIf
   EndIf
   Local $sTestcaseNumber = _ReadWorkbook($g_oWorkbook, Default, $g_sCellTestcase)
   If StringRegExp($sTestcaseNumber, "\d{1,}") == 1 Then ;Search string for 1 or more digits.
	  $sTestcaseNumber = StringRegExpReplace($sTestcaseNumber, "\D", "") ;Remove all non digits form string.
	  Local $sTestcaseID = _CreateID()
	  _CreateXML($sTestcaseNumber, $sTestcaseID)
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
	  Run("QuRickstart_Temp.exe " & $iAmounOfPgdn, "")
	  If $g_bDebugModes == True Then
		 $oExcel.Run("StartTestDebug")
	  Else
		 $oExcel.Run("StartTest")
	  EndIf
	  If $g_bDeleteTestcase == True Then
		 Local $sTestscasePath = $g_sDefaultFolder & $sTestcaseID & " " & $g_sTestscenarioName & ".xml"
		 Local $iDelete = FileDelete($sTestscasePath)
		 If $iDelete == 0 Then
			MsgBox(0, "QuRickstart error!", "Testcase '" & $sTestcaseID & " " & $g_sTestscenarioName & "' could not be deleted, I don't know why...")
		 EndIf
	  EndIf
   Else
	  MsgBox(1, "QuRickstart error!", "The value in cell " & $g_sCellTestcase & "= '" & $sTestcaseNumber & "' contains no number.")
   EndIf
   _Wait()
EndFunc

;----------------------& Functions &----------------------;
;Input: 		1. Number of the testcase.
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

;Input: 		Nothing.
;Returns: 		ID.
;Description: 	Generates an unique ID based on the date and time.
Func _CreateID()
   Local $aiTijd = StringSplit(_NowTime(), ":")
   Local $aiDatum = StringSplit(_NowDate(), "-")
   Local $sUniqueID = "zz" & StringRight($aiDatum[3],2) & $aiDatum[2] & $aiDatum[1] & $aiTijd[1] & $aiTijd[2] & $aiTijd[3]
   Return($sUniqueID)
EndFunc

;Input: 		1. Excel workbook object.
;				2. Excel sheet number.
;				3. Excel cell number.
;Returns: 		Value of a Excel cell.
;Description: 	Reads an Excel workbook -> Sheet -> Cell to return the value.
Func _ReadWorkbook($oWorkbook, $sSheet, $sExcelCell)
   Local $sCellValue = _Excel_RangeRead($oWorkbook, $sSheet, $sExcelCell)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  _Wait()
   EndIf
   Return($sCellValue)
EndFunc

;Input: 		A directory path.
;Returns: 		A number.
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

;Input: 		Nothing.
;Returns: 		Nothing.
;Description: 	Just sleeps forever until a hotkey is pressed.
Func _Wait()
   While True
	  Sleep(100)
   WEnd
EndFunc

;Input: 		Nothing.
;Returns: 		Nothing.
;Description: 	Sets the variable $g_bDebugModes to True and starts the main script.
Func _DebugModes()
   $g_bDebugModes = True
   TrayTip("QuRickstart", "Starting testcase in debug mode!",10,1)
   _Main()
EndFunc

;Input: 		Nothing.
;Returns: 		Nothing.
;Description:	Sets the variable $g_bDebugModes to False and starts the main script.
Func _StartTestNormal()
   $g_bDebugModes = False
   TrayTip("QuRickstart", "Starting testcase!",10,1)
   _Main()
EndFunc

;Input: 		Nothing.
;Returns: 		Nothing.
;Description: 	Exits the script.
Func _ExitScript()
   TrayTip("QuRickstart", "Stopping QuRickstart.... >:(",10,1)
   Sleep(1000)
   Exit
EndFunc

;Author: Rick Ensink. Rick.Ensink@(uwv.nl/sogeti.com), 0645037152