#include <Array.au3>
#include <Excel.au3>
#include <Date.au3>
#include <File.au3>
FileInstall("V:\TSC\K&S\WG Portaal\Rick\GO_eWG_Testscript\AutoIT\QuRickstart\QuRickstart_v8.0\QuRickstart_Temp.exe", @ScriptDir & "/QuRickstart_Temp.exe", 1)

;----------------------& .Ini global variables &----------------------;
Global $g_sIniFilePath 			= @ScriptDir & "\QuRickstart.ini"
Global $g_sHotKeyStart 			= IniRead($g_sIniFilePath, "Hotkeys"	, "HotKey_Start_Normal"			, "F4")
Global $g_sHotKeyDebug 			= IniRead($g_sIniFilePath, "Hotkeys"	, "HotKey_Start_Debug"			, "F3")
Global $g_sHotKeyStop 			= IniRead($g_sIniFilePath, "Hotkeys"	, "HotKey_Stop_QuRickstart"		, "Escape")
Global $g_sWAFWorkbookPath		= IniRead($g_sIniFilePath, "WAF"		, "WAF_Workbook_Path"			, "")
Global $g_bDeleteTestcase 		= IniRead($g_sIniFilePath, "Testcase"	, "Delete_XML_Testcase"			, True)
Global $g_sCellTestcase 		= IniRead($g_sIniFilePath, "Testcase"	, "Excel_Cell_Testcase"			, "B1")
Global $g_sCellTestscenario 	= IniRead($g_sIniFilePath, "CoverSheet"	, "Excel_Cell_Testscenario"		, "B4")
Global $g_sCellDefaultFolder 	= IniRead($g_sIniFilePath, "CoverSheet"	, "Excel_Cell_DefaultFolder"	, "B9")

;----------------------& Other global variables &----------------------;
Global $g_bDebugModes 			= False
Global $g_oWorkbook 			= Null
Global $g_sTitle 				= ""
Global $g_sDefaultFolder 		= ""
Global $g_sTestscenarioName 	= ""

;------------------------& Script &-----------------------;
HotKeySet("{" & $g_sHotKeyStart	& "}", "_StartTestNormal") ;Set Hotkeys and _Wait for input.
HotKeySet("{" & $g_sHotKeyDebug	& "}", "_DebugModes")
HotKeySet("{" & $g_sHotKeyStop 	& "}", "_ExitScript")
_Wait()

Func _Main()
   If $g_sTitle == WinGetTitle("[ACTIVE]") And StringRegExp($g_sTitle, "(?i).*Excel.*") == 1 Then
   Else
	  If StringRegExp(WinGetTitle("[ACTIVE]"), "(?i).*Excel.*") == 1 Then
		 $g_sTitle = WinGetTitle("[ACTIVE]")
		 local $sFilename = StringRight($g_sTitle,StringLen($g_sTitle)-18) ;Removes the text "Microsoft Excel - " from the title.
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
   If StringRegExp($sTestcaseNumber, "\d{1,}") == 1 Then ;Check if string contains 1 or more digits.
	  $sTestcaseNumber = StringRegExpReplace($sTestcaseNumber, "\D", "") ;Remove all non digits from String.
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
   Else ;StringRegExp($sTestcaseNumber, "\d{1,}") == 0 Then.
	  MsgBox(1, "QuRickstart error!", "The value in cell " & $g_sCellTestcase & "= '" & $sTestcaseNumber & "' contains no number.")
   EndIf
   _Wait()
EndFunc

;----------------------& Functions &----------------------;
Func _CreateXML($sTestcaseNumber, $sTestcaseID) ;Navigate in Excel to generate xml.
   Send("!y") ;Navigate Excel menu to validate all tabs.
   Send("{y Down}{5}")
   Sleep(100) ;Siesta needed, otherwise the sript is too fast boi.
   local $sCurrentWindow = WinGetTitle("[ACTIVE]")
   If $sCurrentWindow == "Error Log" Then ;Check if an error is present.
	  _Wait()
   EndIf
   Send("{ENTER}")
   Send("!y") ;Navigate Excel menu to generate xml of current tab.
   Send("{y Down}{8}{TAB 5}")
   Local $iNumberInList = -1 + $sTestcaseNumber
   Send("{DOWN " & $iNumberInList & "}")
   Send("{SPACE}{TAB 6}{SPACE}{TAB}{ENTER}{TAB 2}{LEFT}")
   Send($sTestcaseID & " ")
   Send("{TAB}{ENTER 2}")
EndFunc

Func _CreateID() ;Takes current date and time to create unique ID.
   Local $aiTijd = StringSplit(_NowTime(), ":")
   Local $aiDatum = StringSplit(_NowDate(), "-")
   Local $sUniqueID = "zz" & StringRight($aiDatum[3],2) & $aiDatum[2] & $aiDatum[1] & $aiTijd[1] & $aiTijd[2] & $aiTijd[3] ;YearMonthDay-HourMinutesSeconds
   Return($sUniqueID)
EndFunc

Func _ReadWorkbook($oWorkbook, $sSheet, $sExcelCell) ;Read from workbook. Choose the open sheet, specific cell.
   Local $sCellValue = _Excel_RangeRead($oWorkbook, $sSheet, $sExcelCell)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  _Wait()
   EndIf
   Return($sCellValue)
EndFunc

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

Func _Wait() ;_Waits until a hotkey is pressed.
   While True
	  Sleep(100)
   WEnd
EndFunc

Func _ExitScript() ;Exit script if hotkey is pressed.
   MsgBox(0, "QuRickstart", "Script terminated, siësta initiated!")
   Exit
EndFunc

Func _DebugModes() ;Starts the testscript in debug mode.
   $g_bDebugModes = True
   _Main()
EndFunc

Func _StartTestNormal() ;Starts the testscript without debug mode.
   $g_bDebugModes = False
   _Main()
EndFunc

