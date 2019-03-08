#include <Array.au3>
#include <Excel.au3>
#include <Date.au3>
#include <File.au3>

;----------------------& Variables &----------------------;
Global $g_sWorkbookPadWAF = "V:\TSC\K&S\WG Portaal\Rick\GO_eWG_Testscript\WAF-Execution_v2.0 - eWG_Testscript.xls" ;Path to the WAF excelsheet. Change this when needed.
Global $g_bDeleteTestcase = True
Global $g_sHotKeyStart = "{F4}" ;Hotkey for starting the script.
Global $g_sHotKeyDebug = "{F3}" ;Hotkey for starting the script in debug mode.
Global $g_sHotKeyStop = "{Escape}" ;Hotkey for stoppping the script.
Global $g_bDebugModes = False ;If debug modes is true, start test in debug mode. If false start the test in normal mode.
Global $g_oWorkbook = Null
Global $g_sTitle = ""
Global $g_sCellTestcase = "B1" ;Cell in Excel where to find a number, should be the same in all testscript templates.
Global $g_sDefaultFolder = ""
Global $g_sCellDefaultFolder = "B9"
Global $g_sTestscenarioName = ""
Global $g_sCellTestscenario = "B4"
Global $g_iAmounOfPgdn = 0

;------------------------& Script &-----------------------;
HotKeySet($g_sHotKeyStart, "StartTestNormal") ;Set Hotkeys and wait for input.
HotKeySet($g_sHotKeyDebug, "DebugModes")
HotKeySet($g_sHotKeyStop, "ExitScript")
Wait()

Func Main()
   If $g_sTitle == WinGetTitle("[ACTIVE]") Then
   Else
	  $g_sTitle = WinGetTitle("[ACTIVE]")
	  If StringRegExp($g_sTitle, "(?i).*Excel.*") == 1 Then
		 local $sFilename = StringRight($g_sTitle,StringLen($g_sTitle)-18) ;Removes the text "Microsoft Excel - " from the title.
		 Local $asWorkbookList = ListExcelWorkbooks()
		 For $i = 0 To UBound($asWorkbookList) -1
			If StringRegExp($asWorkbookList[$i][1], $sFilename & ".*") == 1 Then
			   Local $sFilePath = $asWorkbookList[$i][2] & "\" & $asWorkbookList[$i][1]
			   $g_oWorkbook = AttachWorkbook($sFilePath)
			   $g_sDefaultFolder = ReadWorkbook($g_oWorkbook, 1, $g_sCellDefaultFolder)
			   $g_sTestscenarioName = ReadWorkbook($g_oWorkbook, 1, $g_sCellTestscenario)
			   ExitLoop
			Else ;StringRegExp($asWorkbookList[$i][1], $sFilename & ".*") == 1 and continue loop.
			EndIf
		 Next
	  Else ;StringRegExp($g_sTitle, "(?i).*Excel.*") == 0 Then.
		 MsgBox(0, "QuRickstart error!", "Please start the script in an active testscript excelsheet, pannenkoek.")
		 Wait()
	  EndIf
   EndIf
   Local $sTestcaseNumber = ReadWorkbook($g_oWorkbook, Default, $g_sCellTestcase)
   If StringRegExp($sTestcaseNumber, "\d{1,}") == 1 Then ;Check if string contains 1 or more digits.
	  $sTestcaseNumber = StringRegExpReplace($sTestcaseNumber, "\D", "") ;Remove all non digits from String.
	  Local $sTestcaseID = CreateID()
	  CreateXML($sTestcaseNumber, $sTestcaseID)
	  Local $oExcel = OpenExcel()
	  Local $oWorkbookWAF = OpenWorkbook($oExcel,$g_sWorkbookPadWAF)
	  CountFilesDir($g_sDefaultFolder)
	  Run("QuRickstart_StartTest_v2.0.exe " & $g_iAmounOfPgdn, "")
	  If $g_bDebugModes == True Then
		 $oExcel.Run("StartTestDebug")
	  Else ;$g_bDebugModes == True
		 $oExcel.Run("StartTest")
	  EndIf
	  If $g_bDeleteTestcase == True Then
		 Local $sTestscasePath = $g_sDefaultFolder & $sTestcaseID & " " & $g_sTestscenarioName & ".xml"
		 Local $iDelete = FileDelete($sTestscasePath)
		 If $iDelete == 1 Then
		 Else
			MsgBox(0, "QuRickstart error!", "Testcase could not be deleted, just because.")
		 EndIf
	  Else
		 Wait()
	  EndIf
   Else ;StringRegExp($sTestcaseNumber, "\d{1,}") == 0 Then.
	  MsgBox(1, "QuRickstart error!", "The value in cell " & $g_sCellTestcase & "= '" & $sTestcaseNumber & "' contains no number.")
	  Wait()
   EndIf
EndFunc

;----------------------& Functions &----------------------;
Func CreateXML($sTestcaseNumber, $sTestcaseID) ;Navigate in Excel to generate xml.
   Send("!y") ;Navigate Excel menu to validate all tabs.
   Send("{y Down}{5}")
   Sleep(100) ;Siesta needed, otherwise the sript is too fast boi.
   local $sCurrentWindow = WinGetTitle("[ACTIVE]")
   If $sCurrentWindow <> "Error Log" Then ;Check if an error is present.
   Else
	  Wait()
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

Func CreateID() ;Takes current date and time to create unique ID.
   Local $aiTijd = StringSplit(_NowTime(), ":")
   Local $aiDatum = StringSplit(_NowDate(), "-")
   Local $sUniqueID = "zz" & StringRight($aiDatum[3],2) & $aiDatum[2] & $aiDatum[1] & $aiTijd[1] & $aiTijd[2] & $aiTijd[3] ;YearMonthDay-HourMinutesSeconds
   Return($sUniqueID)
EndFunc

Func OpenExcel() ;Start up Excel or connect to a running instance.
   local $oExcel = _Excel_Open()
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error opening excel" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
   Return($oExcel)
EndFunc

Func OpenWorkbook($oExcel, $vWorkbookPad) ;Opens an existing Workbook.
   Local $oWorkbook = _Excel_BookOpen($oExcel, $vWorkbookPad)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error opening the workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
   Return($oWorkbook)
   EndFunc

Func ReadWorkbook($oWorkbook, $sSheet, $sExcelCell) ;Read from workbook. Choose the open sheet, specific cell.
   Local $sCellValue = _Excel_RangeRead($oWorkbook, $sSheet, $sExcelCell)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
   Return($sCellValue)
EndFunc

Func AttachWorkbook($sFilePath) ;Attach to open workbook based on title.
   Local $oWorkbook = _Excel_BookAttach($sFilePath)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error attaching to workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
   Return($oWorkbook)
   EndFunc

Func ListExcelWorkbooks() ;Lists and returns a list of all open Excel workbooks.
   Local $asWorkbookList = _Excel_BookList()
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error listing Excel workbooks." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
   Return($asWorkbookList)
EndFunc

Func CountFilesDir($sDir)
   Local $aFileList = _FileListToArray($sDir)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error opening excel" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
   Local $iFileCount = $aFileList[0]
   If $iFileCount <= 19 Then
	  $g_iAmounOfPgdn = 1
   Else
	  $g_iAmounOfPgdn = (Int(($iFileCount - 19) / 18) + 2)
   EndIf
   ;Return $g_iAmounOfPgdn
EndFunc

Func Wait() ;Waits until a hotkey is pressed.
   While True
	  Sleep(100)
   WEnd
EndFunc

Func ExitScript() ;Exit script if hotkey is pressed.
   MsgBox(0, "QuRickstart", "Script terminated, siësta initiated!")
   Exit
EndFunc

Func DebugModes() ;Starts the testscript in debug mode.
   $g_bDebugModes = True
   Main()
EndFunc

Func StartTestNormal() ;Starts the testscript without debug mode.
   $g_bDebugModes = False
   Main()
EndFunc

