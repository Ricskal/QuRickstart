[Summary:]
This script quickly generates a tescase xml and starts it in the WAF workbook. 
By default the generated xml of the testcase is deleted afterwards (Config param: Delete_XML_Testcase = True).

[How to use:]
1. Add the path of the WAF Workbook in the QuRickstartConfig.ini, parameter: WAF_Workbook_Path = *Your path*
2. Start the QuRickstart.exe. (Visible in the system tray)
3. Open your testscript and navigate to the tescase sheet.
4. All your testcases should be numberd starting at 1. examples.:
	- Test case title 1
	- 2_testcase_title
	- T_3_Testcase
	- 4
	- This is my 5th testcase
5. Press the shortkey to start the script.

Note: The first time you run the script in a testcript it starts up slowly, the second time you use it in the same testscript it is a lot faster.

[Hotkeys:]
The script runs in the background until certain hotkeys are pressed, default:
F3 	= Start test in debug mode	(Config param: HotKey_Start_Debug).
F4 	= start test 			(Config param: HotKey_Start_Normal).
Escape 	= Exit script 			(Config param: HotKey_Stop_QuRickstart).

Want different hotkeys?
Available hokeys: https://www.autoitscript.com/autoit3/docs/functions/HotKeySet.htm (Google "AutoIT HotKeySet")
Syntax for keys: https://www.autoitscript.com/autoit3/docs/functions/Send.htm (Google "AutoIT Send Key")

[Config:]
The config parameters below should not be changed unless the layout of the testscript changes:
Excel_Cell_Testcase 		= B1 	(Cell where to find the testcase name. eg.: Testcase_1.)
Excel_Cell_Testscenario 	= B4 	(Coversheet, cell where to find the testscenario name. eg.: Regression Verzuimmelder.)
Excel_Cell_DefaultFolder 	= B9 	(Coversheet, cell where to find the default folder where the generated testcase xml's are stored.)

The QuRickstartConfig.ini file must be in the same directory as the QuRickstart.exe.

[Other:]
Starting the QuRickstart.exe generates a new .exe file: QuRickstart_Temp.exe. If you want to move the script to another location it's not necessary to move this. 

This is an AutoIT script.