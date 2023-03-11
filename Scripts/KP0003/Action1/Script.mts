Dim dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult

Call spLoadLibrary()
Call spInitiateData("Excel_Report.xlsx", "KP0003.xlsx", "KP0003")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult, "")

If dt_TCID = "TCSN001" Then
	Call Login()
	Call Menu()
	Call Filter_Tipe_Pembayaran("Filter")
ElseIf dt_TCID = "TCSN002" Then
	Call Filter_Tipe_Pembayaran("Reset")
	Call Logout()
End If

Call spReportSave()
REM ========== SUB LOAD LIBRARY
Sub spLoadLibrary()
	Dim objSysInfo, Path_Env, LibFunction
	Set objSysInfo 		= Createobject("Wscript.Network")	
	Path_Env = Environment.Value("Path_Folder")
	LibFunction = Path_Env & "Libraries\"
	LibRepo = Path_Env & "Repositories\"
	LoadFunctionLibrary (LibFunction & "Lib_Report.vbs")
	LoadFunctionLibrary (LibFunction & "Lib_GlobalFunction.qfl")
	
	LoadFunctionLibrary (LibFunction & "Lib_Main.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_SemuaStatus_RawatJalan.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Main.tsr")
End Sub

Sub spGetDatatable()
	REM ---------- Report Data
	dt_TCID					= DataTable.Value("TC_ID", dtLocalSheet)
	dt_ScenarioName		= DataTable.Value("SCENARIO_NAME", dtLocalSheet)
	dt_TestCase				= DataTable.Value("TEST_CASE", dtLocalSheet)
	dt_ExpectedResult		= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
End Sub
