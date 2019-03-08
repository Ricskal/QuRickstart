#include <Array.au3>
#include <Excel.au3>
#include <Date.au3>

;Todo: count files for page down amount. Parameter file + exe. mouse position is kut.m

;----------------------& Variables &----------------------;
Global $g_sWorkbookPadWAF = "V:\TSC\K&S\WG Portaal\Rick\GO_eWG_Testscript\WAF-Execution_v2.0 - eWG_Testscript.xls" ;Path to the WAF excelsheet. Change this when needed.
Global $g_sExcelCell = "B1" ;Cell in Excel where to find a number, should be the same in all testscript templates.
Global $g_iAmounOfPgdn = 50 ;Amount of time to press the button 'PgDn' (Page down) in the WAF tool. Increase if you have to many testscenarios, or just delete them.
Global $g_sHotKeyStart = "{F4}" ;Hotkey for starting the script.
Global $g_sHotKeyDebug = "{F3}" ;Hotkey for starting the script in debug mode.
Global $g_sHotKeyStop = "{Escape}" ;Hotkey for stoppping the script.
Global $g_bDebugModes = False ;If debug modes is true, start test in debug mode. If false start the test in normal mode.
Global $g_oWorkbook = Null
Global $g_sTitle = ""

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
			   ExitLoop
			Else ;StringRegExp($asWorkbookList[$i][1], $sFilename & ".*") == 1 and continue loop.
			EndIf
		 Next
	  Else ;StringRegExp($g_sTitle, "(?i).*Excel.*") == 0 Then.
		 MsgBox(0, "QuRickstart error!", "Please start the script in an active excelsheet pannenkoek.")
		 Wait()
	  EndIf
   EndIf
   Local $sResult = ReadWorkbook($g_oWorkbook, $g_sExcelCell)
   If StringRegExp($sResult, "\d{1,}") == 1 Then ;Check if string contains 1 or more digits.
	  $sResult = StringRegExpReplace($sResult, "\D", "") ;Remove all non digits from String.
	  CreateXML($sResult)
	  OpenWorkbook(OpenExcel(),$g_sWorkbookPadWAF)
	  WinSetState("[ACTIVE]", "", @SW_MAXIMIZE)
	  StartTest()
	  Wait()
   Else ;StringRegExp($sResult, "\d{1,}") == 0 Then.
	  MsgBox(1, "QuRickstart error!", "The value in cell " & $g_sExcelCell & "= '" & $sResult & "' contains no number.")
	  Wait()
   EndIf
EndFunc

;----------------------& Functions &----------------------;
Func CreateXML($sResult) ;Navigate in Excel to generate xml.
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
   Local $iNumberInList = -1 + $sResult
   Send("{DOWN " & $iNumberInList & "}")
   Send("{SPACE}{TAB 6}{SPACE}{TAB}{ENTER}{TAB 2}{LEFT}")
   Send(CreateID() & " ") ;Call Function CreateID for an unique ID.
   Send("{TAB}{ENTER 2}")
EndFunc

Func CreateID() ;Takes current date and time to create unique ID.
   Local $aiTijd = StringSplit(_NowTime(), ":")
   Local $aiDatum = StringSplit(_NowDate(), "-")
   Local $sUniqueID = "(" & $aiDatum[3] & $aiDatum[2] & $aiDatum[1] & "-" & $aiTijd[1] & $aiTijd[2] & $aiTijd[3] & ")" ;YearMonthDay-HourMinutesSeconds
   Return($sUniqueID)
EndFunc

Func StartTest() ;Runs the latest test in WAF excelsheet.
   If $g_bDebugModes == False Then
	  MouseClick($MOUSE_CLICK_LEFT, 1150, 230, 1, 0) ;Move mouse to "Start test".
   Else ;$g_bDebugModes == True
	  MouseClick($MOUSE_CLICK_LEFT, 1150, 340, 1, 0) ;Move mouse to "Start test (Debug modes)".
   EndIf
   Send("{TAB 2}{PGDN " & $g_iAmounOfPgdn & "}{SPACE}{TAB 2}{SPACE}{TAB 4}{SPACE}") ;Navigate the "Start test" menu, and select latest script.
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
   Local $g_oWorkbook = _Excel_BookOpen($oExcel, $vWorkbookPad)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error opening the workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
	  Return($g_oWorkbook)
   EndFunc

Func ReadWorkbook($oWorkbook, $g_sExcelCell) ;Read from workbook. Choose the open sheet, specific cell.
   Local $sResult = _Excel_RangeRead($oWorkbook, Default, $g_sExcelCell)
   If @error Then
	  MsgBox(0, "QuRickstart error!", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Wait()
   EndIf
	  Return($sResult)
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

Func Wait() ;Waits until a hotkey is pressed.
   While 1
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