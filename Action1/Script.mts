On Error Resume Next
Call InitiateGlobalDataTable()
Call InitializeUIXLS(true)

If runBoolean = True Then
	Call executeAutomationTestScriptFromFolder
End If

Call fn_close_all_excel_files()
If err.Number <> 0 Then 
 'MsgBox Err.Description
End If

Function InitiateGlobalDataTable() 
	browserVersion = Trim(LCase(DataTable.Value("Platform","Global")))
	outputDriveDetails = Trim(DataTable.Value("Output_Drive_Details", "Global"))
	globalWaitTime = Trim(DataTable.Value("Wait_Time","Global"))
	globalProjectName = Trim(DataTable.Value("Project_Name", "Global"))
	globalSmsPath = Trim(Environment.value("SMS_Path_File"))
	globalWindowTitle = "Automation_HybridFramework_UFT"
	globalPDF = Trim(Environment.value("PDFTOTEXT_PATH"))
	AS400_Maker_ID = Trim(Environment.value("AS400_Maker_ID"))
	AS400_Maker_Pass = Trim(Environment.value("AS400_Maker_Pass"))
	globalWsPath = Trim(Environment.value("AS400_PATH"))
	globalREPORT = Trim(Environment.value("REPORT_PATH"))
	globalMessageChecker = Trim(Environment.value("MessageChecker"))
	globalBatch = Trim(DataTable.Value("Batch","Global"))
	globalAppName = Trim(DataTable.Value("Package_Name","Global"))
	globalCloseBrowser = Trim(Environment.value("Close_Browser"))
	globalScreenshot = Trim(Environment.value("SCREENSHOT"))
	globalVariable = "START"
	globalPDFTranID = ""
	globalif = -1
	globalPDFextract = "START"
	globalRPT = "START"						  
	Call ResetGlobalVariables
End Function

Function executeAutomationTestScriptFromFolder
	Dim sFileList, currFileName
	Set fso=createobject("Scripting.FileSystemObject")
	iSheet = 1	
	For pos = 0 to Cint(objTree.Nodes.Count.toString) - 1
		If StrComp(objTree.Nodes.item(pos).Checked.toString, "True") = 0 Then
			If Lcase(Trim(DataTable.Value("Batch","Global"))) = "yes" Then
				currFilePath = objTree.Nodes.item(pos).text
				Set objFSO = createobject("Scripting.FileSystemObject")
				Set objFile1 = objFSO.OpenTextFile(currFilePath)
				vFile = objFso.GetFile(currFilePath)
            	batchName = objFso.GetFileName(vFile)
            	batchName = Replace(batchName, ".txt", "")
				Call InitializeTestReportSettingBatch(batchName)
				While objFile1.AtEndOfStream <> true
					currFilePath = Trim(objFile1.ReadLine)
					vFile = fso.GetFile(currFilePath)
            		currFileName = fso.GetFileName(vFile)
            		currFileName = Replace(Replace(currFileName,".xlsx",""),"xls","")
					Call executeAutomationTestScriptFromFile(currFileName , vFile, iSheet, batchName)
					iSheet = iSheet + 1
					Call ResetGlobalVariables	
				Wend
				Call EndBatchReport(batchName)	
				Set objFSO = nothing
				Set objFile1 = nothing
			Else
				currFilePath = objTree.Nodes.item(pos).text
				vFile = fso.GetFile(currFilePath)
            	currFileName = fso.GetFileName(vFile)
            	currFileName = Replace(Replace(currFileName,".xlsx",""),"xls","")
				Call executeAutomationTestScriptFromFile(currFileName , vFile, iSheet, batchName)
				iSheet = iSheet + 1	
				Call ResetGlobalVariables
			End If			
		End If
	Next
	Set fso=nothing
End Function

Function executeAutomationTestScriptFromFile(fileName, vFile, iSheet, batchName)
	On Error Resume Next
	Dim ExcelApp, ExcelFile, sheetTestData, sheetSourceName, sheetDestinationName
	sheetSourceName = "TestScript"
	sheetTestData = "TestData"
	sheetExecutionName = sheetDestinationName
	Set ExcelApp = CreateObject("Excel.Application")
	Set ExcelFile = ExcelApp.WorkBooks.Open(vFile, 0, True)
	'ExcelFile.Visible = false 'For windows 10 only
	Set ExcelTestScriptSheet = ExcelApp.WorkSheets(sheetSourceName)

	If isTestDataExist(sheetTestData, ExcelApp.WorkSheets) Then
		Dim testDataRowCount, testDataRowIndex, ExcelTestDataSheet, reportFileName
		Set ExcelTestDataSheet = ExcelApp.WorkSheets(sheetTestData)
		testDataRowCount = ExcelTestDataSheet.UsedRange.rows.count
		For testDataRowIndex = 2 to testDataRowCount
			If ExcelTestDataSheet.Cells(testDataRowIndex, 1) = "Y" Then
				sheetDestinationName = "Action" & iSheet
				reportFileName = ExcelTestDataSheet.Cells(testDataRowIndex, 2)
				 	Set qSheet = DataTable.GetSheet(sheetDestinationName)
					If Err.Number<>0 Then
						Set qSheet = DataTable.AddSheet(sheetDestinationName)
						Err.Clear
					End If
					If Lcase(Environment.value("GET_NOW")) = "true" Then
						strAfter = Now
					Else
						strAfter = getNow()
					End If 
					DataTable.ImportSheet vFile,sheetSourceName,sheetDestinationName
					Call ImportTSwithTestData(sheetDestinationName, ExcelTestScriptSheet, ExcelTestDataSheet, testDataRowIndex)
				
				End If
				'ExcelFile.Close False 
				'ExcelApp.Quit
				Call executeAndGenerateReport(reportFileName, sheetDestinationName, batchName)
				Call ResetGlobalVariables
			End If
		Next
	Else
		sheetDestinationName = "Action" & iSheet
		If Lcase(Environment.value("IMPORT_SHEET")) = "true" Then
			colCount = ExcelTestScriptSheet.UsedRange.Columns.Count	
			For sColumnindex = 1 to colCount
				sRowValue = ExcelTestScriptSheet.Cells(1, sColumnIndex).value
				ExcelTestScriptSheet.Cells(1, sColumnIndex).value = Replace(sRowValue, " ","_")
			Next 
			Set qSheet = DataTable.GetSheet(sheetDestinationName)
			If Err.Number<>0 Then
				Set qSheet = DataTable.AddSheet(sheetDestinationName)
				Err.Clear
			End If
			DataTable.ImportSheet vFile,sheetSourceName,sheetDestinationName 
		Else
			Call ImportSheetFromXLSX(ExcelTestScriptSheet, sheetDestinationName)
		End If
		If Lcase(Environment.value("GET_NOW")) = "true" Then
			strAfter = Now
		Else
			strAfter = getNow()
		End If	
		ExcelFile.Close False 
		ExcelApp.Quit		
		Call executeAndGenerateReport(fileName, sheetDestinationName, batchName)
	End If
 
 	ExcelFile.Close False 
	ExcelFile.DisplayAlerts = False
	ExcelApp.Quit
	Set ExcelApp = Nothing
	Set ExcelFile = Nothing
 End Function

Function isTestDataExist(testDataWorkSheetName, Worksheets)
	For i = 1 To Worksheets.Count
		If Worksheets(i).Name = testDataWorkSheetName Then
			isTestDataExist = True
		End If
	Next
End Function

Function ResetGlobalVariables
	globalAppObjHwnd = 0
	glbalAppObjVersion = ""
	glbalDeviceModel= ""
	glbalDeviceUDID = ""
	glbalDeviceVersion= ""
	globalAppObjFind = False
	globalBrowserOpen = False
	globalFailed = 0
	imgIndex = 0
	strIndex = 0
	globalMessageChecker = ""
	glbalReferenceNo = ""
	glbalIndicator = ""
	Environment.value("T_REF").RemoveAll
	globalif = -1
End Function
 @@ hightlight id_;_8641_;_script infofile_;_ZIP::ssf63.xml_;_
Function executeAndGenerateReport(fileName, sheetDestinationName, batchName)
	Dim strIndex, imgIndex
	Call InitializeTestReportSetting(fileName,globalMessageChecker)
	If globalHashIteration.contains(fileName) = "True" Then
		testValue = cint(globalHashIteration.Item (fileName))
		testValue = testValue + 1
		globalHashIteration.Item (fileName) = cstr(testValue)
	Else	
		globalHashIteration.add fileName, "1"
	End If
	If findBrowserType() = ".*" Then

	ElseIf Instr(browserVersion, "window") > 0 Then
	ElseIf Trim(LCase(browserVersion)) = "ie" Or Trim(LCase(browserVersion)) = "chrome" Or Trim(LCase(browserVersion)) = "firefox" or browserVersion = "edge" Or browserVersion = "safari" Then
		'Close_Browser() 'First Step to Close All Browser before run test script!
	End If
	For i = 1 To DataTable.GetSheet(sheetDestinationName).GetRowCount Step 1
		DataTable.SetCurrentRow(i)
Call executeStepExcel(fileName, sheetDestinationName, i)	
		If err.Number <> 0 Then 
			MsgBox Err.Description
			Call AddToTestResult(fileName,"Field level error at " & Trim(DataTable.Value("Step_Name",sheetDestinationName)) , Err.Description ,"","-1", "FAIL", globalMessageChecker, "") 
		End If 		
	Next
	Call Func_ReportingEventsExternal(fileName,globalMessageChecker,browserVersion,batchName)
End Function

 Function fn_close_all_excel_files()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer &"\root\cimv2")
	
	Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process Where Name = 'EXCEL.EXE'")
	For Each objProcess in colProcessList @@ hightlight id_;_15338232_;_script infofile_;_ZIP::ssf1010.xml_;_
		objProcess.terminate()
	Next
	
	Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'WINWORD.EXE'")
	For Each objProcess in colProcessList
	    objProcess.Terminate()
	Next 
	
	Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process Where Name = 'ONENOTE.EXE'")
	For Each objProcess in colProcessList
		objProcess.terminate()
	Next
	
	Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process Where Name = 'ONENOTEM.EXE'")
	For Each objProcess in colProcessList @@ hightlight id_;_984918_;_script infofile_;_ZIP::ssf962.xml_;_
		objProcess.terminate()
	Next
End Function

