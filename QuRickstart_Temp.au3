;----------------------& Variables &----------------------;
Global $g_iAmounOfPgdn

;------------------------& Script &-----------------------;
If $CmdLine[0] = 0 Then
     MsgBox(0, "QuRickstart error!", "You should never see this error.")
	 Exit
Else
   $g_iAmounOfPgdn = $CmdLine[1] ; assign the passed parameter to $g_iAmounOfPgdn.
EndIf
For $i = 0 To 30000 Step 1
   Local $sCurrentWindowTitle = WinGetTitle("[Active]")
   If $sCurrentWindowTitle == "Select Testscenario's" Then
	  Send("{TAB 2}{PGDN " & $g_iAmounOfPgdn & "}{SPACE}{TAB 2}{SPACE}{TAB 4}{SPACE}")
	  Exit
   EndIf
Next



