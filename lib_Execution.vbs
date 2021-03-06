
Sub EXEC_ModuleDriver( UNIVERSAL)

        Dim  OBJ, sEnv, sExecController, sModule, sSubModule, sTestDataBase , oFSO , oF , fileContent
'        Set UNIVERSAL = CreateObject ("Scripting.Dictionary")
'        Set EXEC = CreateObject("Scripting.Dictionary")	    
        EXEC.RemoveAll
        UNIVERSAL.RemoveAll		           
        subCreateExecutionResultFolder cTESTRESULTLOGFOLDERPATH, EXEC '''Creating the top level html structure
        sExecController = mController
        sEnv ="TestSet1"
        

        DR_SUITE EXEC, UNIVERSAL, sEnv, sExecController        
        Set oFSO  = createobject("scripting.Filesystemobject")
        
        If oFSO.FileExists(Exec.Item("ExecutionResultHtml")) Then
            Set oF = oFSO.OpenTextFile(Exec.Item("ExecutionResultHtml"))
            fileContent = oF.ReadAll
            eEnd=formatDateTime(Now)
            fileContent = Replace(fileContent , "eEnd" , cstr(eEnd) )
            fileContent = Replace(fileContent , "etotalTime" ,  SplitSec(DateDiff("s",eStart,eEnd)))
            
            fileContent = Replace(fileContent , "totalTX" , cstr(iPassCountSuite+iFailCountSuite))
            fileContent = Replace(fileContent , "totalPX" , cstr(iPassCountSuite))
            fileContent = Replace(fileContent , "totalFX" , cstr(iFailCountSuite))
            
            oF.Close
            Set oF = nothing            
            Set oF = oFSO.OpenTextFile(Exec.Item("ExecutionResultHtml") , 2 )
            oF.Write fileContent
            oF.Close
            Set oF = nothing            
        End If

        If oFSO.FileExists(Environment.Value("ExecutionResultHtml")) Then
            systemutil.Run Environment.Value("ExecutionResultHtml")
            Sub_UploadResultoALM(Environment.Value("ExecutionResultHtml") ) 								'Upload Suite test result html file to the Testset in ALM
	    Sub_DisconnectFromALM	
        End If

        Set oFSO = Nothing
End Sub


Function DR_SUITE (EXEC, UNIVERSAL, sEnv, sExecController)

    Dim arrTestCases, iTestCaseCount, arrModule,iModuleCount
    Dim qtpApp ,qtpRepositories
    Dim strName,strOrName,strOrPath , oFSO , oF , fileContent	      
    'Set qtpApp = CreateObject("QuickTest.Application")   
    Dim anyORUploaded
    anyORUploaded=false
   
    Dim ParentObject
    Dim RepositoryFrom    
	Set	OBJ=CreateObject("Scripting.Dictionary")
	OBJ.RemoveAll
	
	'Create Object Rpository and load to OBJ dictionary if the OR Type is "ProjectLevel"
	If ORType="ProjectLevel" Then	
		Dim projectORFilePath
		projectORFilePath=ORFolderPath & "\" & projectOR
		Set qtpRepositories = qtpApp.Test.Actions("Action1").ObjectRepositories 
	
	    qtpRepositories.Add(projectORFilePath)
	    anyORUploaded=true	
		Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")
		RepositoryFrom.Load projectORFilePath
		set allObjectsCollection=RepositoryFrom.GetAllObjects
		getOBJ RepositoryFrom,allObjectsCollection,ParentObject',""',1
		
	End If    
'	getOBJFromFile
    arrModule =fGetProjectModules(sExecController,"tbl_modules")     

    For iModuleCount= 0 to UBound (arrModule)    
        iPassCount = 0
        iFailCount = 0
        Environment.Value("executionFurther") = True
        EXEC.Item("ModuleResultHtml") = subGenerateResultLog_Swap(EXEC, arrModule(iModuleCount,1),arrModule(iModuleCount,0) )
        EXEC.Item("MODULEDATAFILE") = arrModule(iModuleCount,2)
                    
        'Create Object Rpository and load to OBJ dictionary if the OR Type is "ModuleWise"
		If ORType="ModuleWise" Then	
		    Dim orPath,sParentObject
		    Dim i,orNames,arrORNames
'			Set qtpRepositories = qtpApp.Test.Actions("Action1").ObjectRepositories 
'			Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")
			 
			orNames = arrModule(iModuleCount,4)
			If (len(orNames)>0) Then
				arrORNames=split (orNames,";")	
				
				
			
			For i = 0 to UBound(arrORNames)		
				Set qtpRepositories = qtpApp.Test.Actions("Action1").ObjectRepositories 
				Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")			
				orPath=ORFolderPath & "\" & arrORNames(i)
				qtpRepositories.Add(orPath)
				anyORUploaded=true
				RepositoryFrom.Load orPath	
				set allObjectsCollection=RepositoryFrom.GetAllObjects ' This is added to rduce to time creating OBJ.item
				getOBJ RepositoryFrom,allObjectsCollection,sParentObject',""',1
				set allObjectsCollection=nothing 'memory leaks
				Set RepositoryFrom =nothing
			Next
			
			
			End If
		
		End If  

        arrTestCases = fGetTestCases( arrModule(iModuleCount,1),arrModule(iModuleCount,0))
        
        For iTestCaseCount = 0 to UBound(arrTestCases)
            UNIVERSAL.add "TESTFLOW", "Following Steps Completed ::"
            EXEC.Item("TCId")=arrTestCases(iTestCaseCount,1)
			arrSheetName = Split(EXEC.Item("TCId"),"_")
            sSheetName = arrSheetName(1)&"_"&arrSheetName(2)
            DR_TESTCASE UNIVERSAL, EXEC, "tbl_testdata", arrTestCases(iTestCaseCount,0), arrTestCases(iTestCaseCount,1), "tbl_testcase", sEnv,iTestCaseCount ,arrTestCases(iTestCaseCount,2) 
            UNIVERSAL.RemoveAll( )
            wait(2)
        Next
         
        Set oFSO  = createobject("scripting.Filesystemobject")
        
        If oFSO.FileExists(Exec.Item("ModuleResultHtml")) Then
            Set oF = oFSO.OpenTextFile(Exec.Item("ModuleResultHtml"))
            fileContent = oF.ReadAll
            mEnd=formatDateTime(Now)
            fileContent = Replace(fileContent , "mEnd" , cstr(mEnd) )
            fileContent = Replace(fileContent , "mtotalTime" ,  SplitSec(DateDiff("s",mStart,mEnd)) )        
      		oF.Close
            oF.Close
            Set oF = nothing
       		Set oF = oFSO.OpenTextFile(Exec.Item("ModuleResultHtml") , 2 )
            oF.Write fileContent
            oF.Close
            Set oF = nothing
         End If

        Set oFSO = Nothing    
        subWriteModuleResult EXEC , arrModule(iModuleCount,0) , arrModule(iModuleCount,3)

       	If ORType="ModuleWise" and anyORUploaded=true Then
			qtpRepositories.RemoveAll
			OBJ.RemoveAll
		End if 	

    Next
    
        If ORType="ProjectLevel" and anyORUploaded=true Then
			qtpRepositories.RemoveAll
			OBJ.RemoveAll
		End if
		


End Function


Function DR_TESTCASE(UNIVERSAL, EXEC, sDataTable, TCID, TCNAME, sResultTable, sEnvironment,TCNo,TCDesription)

    Dim oFSO,oF,fileContent

    subInitialTestSetUp EXEC, UNIVERSAL, EXEC.Item("MODULEDATAFILE"),sDataTable, TCID, sEnvironment
    iStepCount=0
    execute( TCNAME )
    Set oFSO  = createobject("scripting.Filesystemobject")

    If oFSO.FileExists(Environment.Value("TestCaseResultHtmlPath")) Then
        Set oF = oFSO.OpenTextFile(Environment.Value("TestCaseResultHtmlPath") , 1)
        fileContent = oF.ReadAll
        tEnd=formatDateTime(Now)
            fileContent = Replace(fileContent , "tEnd" , cstr(tEnd) )
            fileContent = Replace(fileContent , "ttotalTime" ,  SplitSec(DateDiff("s",tStart,tEnd)) )        
        oF.Close
        Set oF = nothing        
        Set oF = oFSO.OpenTextFile(Environment.Value("TestCaseResultHtmlPath") , 2 )
        oF.Write fileContent
        oF.Close
        Set oF = nothing        
    End If

    subWriteTestCaseResult EXEC , TCNo , TCNAME , TCDesription    

End Function


Sub exportResultInToLocalMachine

		TCNAME = Environment("TestName")
		tcDtName = Replace(TCNAME," ","_") & "_" &  EXEC.Item("DateTime")		
		DataTable.Export EXEC.Item("TCFOLDER") & "\" & tcDtName & ".xls"

End Sub




















'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

' This Function Initializes the Variables governing Functions
Sub subInitializeFunctionVariables()

	 		bKeywordPF 			  = true
			bMethodPF 				= false
			sKeywordError 		= ""
			

End Sub

Sub subInitializeTestCaseVariables()

	 	bTestCasePF 			= False
		sTestCaseMessage  = ""
		sTestCaseError 		  = ""
		bMethodPF 				 = True
		sMethodError 		    = ""
		iFunctionIteration = 1
		bResultIndicator =  True
'		iStepCount =0
        

End Sub

Function fGetValue (sData)
		'Split(sData, ";;",
		On Error Resume Next
		If sData <> Empty  Then
			If Ubound(Split(sData,";;"))>0 Then
						fGetValue = Split(sData,";;")(iFunctionIteration -1)
			Else				
						fGetValue=sData
			End If
		else
				fGetValue =""
		End If
	
End Function

Sub mFillUniversal(UNIVERSAL, sDataBase, sDataTable, TCID)
		On Error Resume Next

		 Dim arrTestData, strConnection, strSQL
		
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cTESTDATAFOLDERPATH & sDataBase & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
		strSQL = "SELECT Desc,Val FROM [" & sDataTable & "$] where TCID = '"+TCID+"'"  
		'msgBox strSQL
		UNIVERSAL.RemoveAll
		arrTestData = mGetMultipleData(strConnection, strSQL) ' Fetch the Test Data for Specific Test case
		For iDataCount = 0 to UBound(arrTestData)
				UNIVERSAL.Add arrTestData(iDataCount, 0), arrTestData(iDataCount, 1)
		Next

End Sub

'===================== sub_FillUniversalData ===================================

' Following Function fetches the test data from Data file and then adds the Keys and Items to the Dictionary object

Sub sub_FillUniversalData(UNIVERSAL, sDataBase, sDataTable, TCID, sENV)

			On Error Resume Next
			UNIVERSAL.RemoveAll
			Dim arrData, strConnection, sSQLQuery, arrDescKeyword, arrDescVal, sKeywords, objWSHShell,iKeyCount
			Set objWSHShell =  CreateObject("WScript.Shell")
			'TCID = "MODULE1_001"
			sKeywords = ""
			
			sFullExcelPath = cTESTDATAFOLDERPATH & sDataBase
			sTestCaseName = TCID
			strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cTESTDATAFOLDERPATH & sDataBase & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
			sSQLQuery = "SELECT Desc,"& sDataSetColumn &" FROM [" & sDataTable & "$] where TCID = '"+TCID+"'"  
			arrData = mGetMultipleData(strConnection, sSQLQuery)
			
			iKeyCount = -1
			
			For iCount = 0 to UBound(arrData)
					If Left( arrData(iCount,1), 4) = "KEY_" Then
							iKeyCount = iKeyCount + 1
					End If
			Next
			
			ReDim arrDescKeyword(iKeyCount, 1)
			ReDim  arrDescVal(UBound(arrData) - iKeyCount - 1, 1)
			
			iKeyCount = -1
			iValCount = -1
			
			For iCount = 0 to UBound(arrData)
					If Left( arrData(iCount,1), 4) = "KEY_" Then
								iKeyCount = iKeyCount + 1
								arrDescKeyword(iKeyCount,0) = arrData(iCount,0)
								arrDescKeyword(iKeyCount,1) = arrData(iCount,1)
								sKeywords = sKeywords & "'" & arrData(iCount,1) & "',"
					Else
								iValCount = iValCount + 1
								arrDescVal(iValCount,0) = arrData(iCount,0)
								arrDescVal(iValCount,1) = arrData(iCount,1)
					End If
			Next
			sKeywords = Left(sKeywords, Len(sKeywords) - 1)
			
			strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cTESTDATAFOLDERPATH & "DataRepository.xls" & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
			sSQLQuery = "SELECT Keyword,KeywordVal FROM [" & sENV & "$] where KEYWORD IN (" & sKeywords & ")"
			arrData = mGetMultipleData(strConnection, sSQLQuery)
			
'			For iCount = 0 to UBound(arrData)
'					arrDescKeyword(iCount,1) = arrData(iCount,0)
'			Next

			For iCount = 0 to UBound(arrData)
				For iKeyCount=0 to UBound(arrDescKeyword)
						If  arrDescKeyword(iCount,1) = arrData(iKeyCount,0) Then
	
                                        arrDescKeyword(iCount,1) = arrData(iKeyCount,1)
										Exit For 
						End If
				Next

			Next

			
			For iCount = 0 to UBound(arrDescKeyword)
					UNIVERSAL.Add arrDescKeyword(iCount,0), arrDescKeyword(iCount,1)
			Next
			
			For iCount = 0 to UBound(arrDescVal)
					UNIVERSAL.Add arrDescVal(iCount,0), arrDescVal(iCount,1)
			Next

			arrValuesFromDic = UNIVERSAL.Items
			arrDescFromDic = UNIVERSAL.Keys
			
			
Dim sDesc, sValues
sDesc = ""
			For iCount = 0 to UBound(arrDescFromDic)
					sDesc = sDesc & arrDescFromDic(iCount) & "-----------" & arrValuesFromDic(iCount) & Chr(10)
			Next
			'objWSHShell.Popup sDesc, 5, "Test Data for the Work Flow" 
End Sub

Sub subTestCaseVerification(EXEC, sDataBase, sDataTable, TCID,sTestFlow)

		If bMethodPF = True Then
				subUpdateResultFile UNIVERSAL,EXEC, sDataBase, sDataTable, TCID, "Pass", Left("Test Case Passed " &  Replace(sTestCaseMessage,"""","'"),255),sTestFlow
		Else
				subUpdateResultFile  UNIVERSAL, EXEC, sDataBase, sDataTable, TCID, "Fail",  Left("Test Case Failed " & Replace(sMethodError,"""","'"), 255),sTestFlow

		End If

End Sub
'===================== sub_FillUniversalData ===================================



Sub subInitialTestSetUp(EXEC, UNIVERSAL, sDataBase, sDataTable, TCNAME, sEnvironment)

	On Error Resume Next
	Dim objFSO, filetxt
	EXEC.Item("TCFOLDER") = EXEC.Item("ModuleResultFolder") & "\" & Replace(TCNAME," ", "_")  & "_" &Hour(now) & minute(now)& second(now) &"_" & sDataSetColumn 
	EXEC.Item("TCFOLDERName") = Replace(TCNAME," ", "_")
	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.CreateFolder EXEC.Item("TCFOLDER")
	
	EXEC.Item("sSnapshotFolder") = EXEC.Item("TCFOLDER") & "\Snapshot"
	
	objFSO.CreateFolder(EXEC.Item("sSnapshotFolder"))
	
	EXEC.Item("TestCaseResultHtmlPath") = EXEC.Item("TCFOLDER") & "\" & Replace(TCNAME," ", "_") & ".html" 
	Environment.Value("TestCaseResultHtmlPath") = EXEC.Item("TestCaseResultHtmlPath")
	
	Set filetxt = objFSO.CreateTextFile(EXEC.Item("TestCaseResultHtmlPath") , true)
	
	    'logoPath=cROOTPATH &"Layer_Test\Logo\logo.png"
		strFileHeading="<span style='position:absolute;z-index:3;left:36px;top:1px;width:1 px;height:76px'><table cellpadding=0 cellspacing=0><tr><td width=651 height=76 valign=middle style='vertical-align:top'><![endif]>	  <div v:shape=""_x0000_s1028"" style='padding:0pt 0pt 0pt 0pt' class=shape>  <p class=Heading010 style='margin-bottom:0pt'><span lang=en-US  style='font-size:36.0pt;font-family:""Goudy Old Style"";font-weight:bold;	  language:en-US'><span style='position:absolute;left:5px;top:50px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:60.0pt;font-family:Arial;color:#888888;font-weight:normal;language:en-US'>"+projectName+"</span></p></td></tr></table></span></span></p>	  </div> <![if !vml]></td> </tr></table></span>"
		'compay_logo="<span style='position:absolute;left:1000px;top:30px'><img src="""+logoPath+"""></span>"
	
		strFileHeading= "<span style='position:absolute;left:25px;top:110px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:30pt;font-family:Impact;color:#854442;font-weight:Normal;language:en-US'>Testcase : " + Ucase (Replace(TCNAME , "_" , " " ))+"</span></p></td></tr></table></span>"
		
		strFileHeading2= "<span style='position:absolute;left:30px;top:160px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>THE DETAILS OF THE STEPS IN TESTCASE '" + Ucase (Replace(TCNAME , "_" , " " ))+"' CAN BE SEEN BELOW</span></p></td></tr></table></span>"
		
		strBox= "<span style='position:absolute;left:0px;top:185px'><table border=0 style='border-collapse:collapse'><tr><td width= 2000 height=25  bgcolor= ""#4b3832""></td></tr></table></span>"
		strExecutionDate= "<span style='position:absolute;left:800px;top:190px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:13pt;font-family:Calibri;color:#fff4e6;font-weight:normal;language:en-US'>EXECUTION DATE:</span></p></td></tr></table></span>"
		
		
		strStart= "<span style='position:absolute;left:30px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>START DATE & TIME:</span></p></td></tr></table></span>"
		strEnd= "<span style='position:absolute;left:400px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>END DATE & TIME:</span></p></td></tr></table></span>"
		strTotal= "<span style='position:absolute;left:750px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>TOTAL TIME:</span></p></td></tr></table></span>"
		tStart=formatDateTime(Now)
		strStartValue= "<span style='position:absolute;left:170px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>" +cstr(tStart)+"</span></p></td></tr></table></span>"
		strEndValue= "<span style='position:absolute;left:540px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>tEnd</span></p></td></tr></table></span>"
		strTotalValue= "<span style='position:absolute;left:840px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>ttotalTime</span></p></td></tr></table></span>"
		strBox2= "<span style='position:absolute;left:0px;top:240px'><table border=0 style='border-collapse:collapse'><tr><td width= 2000 height=1  bgcolor= ""#be9b7b""></td></tr></table></span>"
		strTableHeading = "<span style='position:absolute;left:30px;top:300px'><table align = 'centre' border='1' width='' style='border-collapse:collapse ' bordercolor = ""#be9b7b"" cellpadding ='5'><tr><th bgcolor=""#be9b7b"" width='40' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>S.NO.</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>STEP DESCRIPTION</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>ACTUAL RESULT</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>STEP STATUS</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>SNAPSHOT</th></tr>"                                                                  
	
	'	filetxt.WriteLine(compay_logo)
		filetxt.WriteLine(strFileHeading)
		filetxt.WriteLine(strFileHeading2)
		filetxt.WriteLine(strBox)
		filetxt.WriteLine(strStart)
		filetxt.WriteLine(strEnd)
		filetxt.WriteLine(strTotal)
		filetxt.WriteLine(strStartValue)
		filetxt.WriteLine(strEndValue)
		filetxt.WriteLine(strTotalValue)
		filetxt.WriteLine(strBox2)	
		filetxt.WriteLine(strTableHeading)

    Set filesys = Nothing
    Set filetxt = Nothing  		
		
	subInitializeTestCaseVariables()' Initialize Test Case Variables	
	sub_FillUniversalData UNIVERSAL, sDataBase, sDataTable, TCNAME, sEnvironment
		
End Sub



'===================== subInitialTestSetUp ===================================

Sub subKillProcesses(sProcessName)
		SystemUtil.CloseProcessByName(sProcessName)
End Sub

Sub subCreateResultLogFile(EXEC, sTaskFile)

		On Error Resume Next
		Dim objFSO, iCount
		EXEC.Add "MODULECONFIGFILE",  cTESTCONFIGFOLDERPATH & sTaskFile
		EXEC.Add "RESULTFILE", cTESTRESULTLOGFOLDERPATH & Left(sTaskFile, len(sTaskFile)-4) & "_Result_" &mGetTimeStampLong1 & ".xls"

		Set objFSO = CreateObject("Scripting.FileSystemObject")
			objFSO.CopyFile EXEC.Item("MODULECONFIGFILE"), EXEC.Item("RESULTFILE")
		Set objFSO = nothing

End Sub


Sub subMethodVerification(EXEC, sMethodName, sMethodMessge)
						
	On Error Resume Next
	'CaptureSnapshot Browser(Environment("sBrowser")).Page(Environment("sPage")),  EXEC.Item("TCFOLDER") & "\" & sMethodName & "_" & CStr(iFunctionIteration) ' == This function will store the last page screen shot Image
	iFunctionIteration = 1
	If bKeywordPF = True  AND Err.Number = 0 Then
		Reporter.ReportEvent micPass,sMethodName, " Successful  " & sMethodMessge
		bMethodPF = True
	Else
        sMethodError = "Function '" & sMethodName & "' Failed.  " & sKeywordError & "     " & Err.Description
		Reporter.ReportEvent micFail, sMethodName, " Failed  " & sMethodError
		bMethodPF = False
	End If

End Sub


Function fGetTestCases( sDataBase, sDataTable)

		On Error Resume Next
		 Dim strConnection, strSQL
		
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cTESTCONFIGFOLDERPATH & mController & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
'		strSQL = "SELECT TCID, TCName FROM [" & sDataTable & "$] where SubModule = '" & sSubModule & "' and Active = 'Yes'"  


		strSQL = "SELECT TCID, TCName, Description, BU FROM ["&sDataBase&"$] where Active = 'Yes'"
		'strSQL = "SELECT TCID, TCName, Description FROM [" & sDataTable & "$] where Active = 'Yes'"
		
		'msgBox strSQL
		fGetTestCases = mGetMultipleData(strConnection, strSQL) ' Fetch the Test Data for Specific Test case

End Function


Function fGetProjectModules(sProjExeSheet, sTable)
 
		On Error Resume Next
		Dim strConnection, sSQLQuery

		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cTESTCONFIGFOLDERPATH & sProjExeSheet & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
		sSQLQuery = "SELECT ModuleName,ExecutionController,DataSheet,Description,ObjectRpository FROM [" & sTable & "$] where Active = 'Yes'"  

		fGetProjectModules = mGetMultipleData(strConnection, sSQLQuery)

End Function


Sub fLibTempPopup(sTitle, sMessage)
		Dim objShell
		Set objShell = CreateObject("WScript.Shell")
		objShell.Popup sMessage,5,sTitle
End Sub
'=====================  Kedarnath Patil =====================


' ===================== Seetha Viswambharan =========

Dim vBrowserType
vBrowserType = "IE"

Function CaptureSnapshot(vParentObject,vSnapshotName)
	On Error Resume Next
   Dim vHTMLCode,vFso,vFo,vExtension,vURL
   vURL = "https://qa-online.portal.nokiasiemensnetworks.com/"
   If vBrowserType = "IE" Then
		vExtension = ".html"
		vFileName = vSnapshotName & vExtension		
		vHTMLCode =vParentObject.Object.all.tags("html")(0).outerhtml		
		Set vFso = CreateObject("Scripting.FileSystemObject")
		Set vFo = vFso.OpenTextFile(vFileName,2,True,-1)
		vFo.Write "<base href=""" & vURL & """/>"
		vFo.Write vHTMLCode
		vFo.close
	Else
		vExtension = ".png" 
		vFileName = vFilePath & vSnapshotName & vExtension
		vParentObject.CaptureBitmap vFileName
	End If
	
End Function

' ===================== Seetha Viswambharan =========


' ===================== Ritesh Gawde =========

Sub subGenerateResultLog__OLD(vTaskFolder )

	Dim vRootFolder,  objExcel , vFileSysObj, objWorkBook, TaskResultFolder, objWorkSheet
	  
vRootFolder = cROOTPATH

	TimeStamp = date&"_"& Hour(Now)&"."&Minute(Now)&"."&Second(Now)

	Set objExcel = CreateObject("EXCEL.APPLICATION")
	Set vFileSysObj= CreateObject("Scripting.FileSystemObject")
	Set TaskResultFolder= vFileSysObj.CreateFolder(vRootFolder&"TestResultLogs\"&vTaskFolder&"_"&TimeStamp)

	Set objWorkBook = objExcel.Workbooks.Open(vRootFolder &"\Layer_Test\TestConfig\"&vTaskFolder&".xls")

	objWorkBook.SaveAs (TaskResultFolder&"\"&vTaskFolder&"_ResultLog_"&TimeStamp&".xls")

	Set objWorkSheet= objWorkBook.Sheets("tbl_testcase")'(vTaskFolder)

	Col = 1
	While (objWorkSheet.Cells(1, Col).value <> "")
		Col =Col +1
	Wend

	objWorkSheet.Cells(1, Col).value = "Status"
	objWorkSheet.Cells(1, Col+1).value = "TimeStamp"
	objWorkSheet.Cells(1, Col+2).value = "Remark"
	objWorkSheet.Cells(1, Col+3).value = "TestFlow"
	objWorkSheet.Cells(1, Col+4).value = "SnapShots"

	
	 objWorkSheet.Range("A1:H1").Interior.ColorIndex = 15
	 objWorkSheet.Range("A1:H1").Font.Bold=15

	objWorkSheet.Cells.Select
    objWorkSheet.Cells.EntireColumn.AutoFit
    objWorkSheet.Cells.EntireColumn.AutoFit
	objWorkSheet.Columns("E:E").ColumnWidth = 15
    objWorkSheet.Columns("F:F").ColumnWidth =  20
    objWorkSheet.Columns("G:G").ColumnWidth = 20
    objWorkSheet.Columns("H:H").ColumnWidth = 25

	objWorkBook.Save
	objWorkBook.Close True
	Set objWorkBook = Nothing
	Set objExcel = Nothing

End Sub

''' Ratnesh - Start

Sub subCreateExecutionResultFolder(cTESTRESULTLOGFOLDERPATH, EXEC )
		
		Dim filesys,siletxt,sExecutionHtmlFilePath

		If EXEC.Exists("sExecutionResultFolderName") Then
			EXEC.Item("sExecutionResultFolderName") = "Execution_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date) & "_" & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now)
		Else
			EXEC.add "sExecutionResultFolderName" , "Execution_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date) & "_" & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now)
		End If		
		
		If EXEC.Exists("sExecutionResultFolder") Then
			EXEC.Item("sExecutionResultFolder") = cTESTRESULTLOGFOLDERPATH & EXEC.Item("sExecutionResultFolderName")
		Else
			EXEC.add "sExecutionResultFolder" , cTESTRESULTLOGFOLDERPATH & EXEC.Item("sExecutionResultFolderName")
		End If			
		
		If EXEC.Exists("sLogFolder") Then
			EXEC.Item("sLogFolder") = EXEC.Item("sExecutionResultFolder") & "\Logo"
		Else
			EXEC.add "sLogFolder" , EXEC.Item("sExecutionResultFolder") & "\Logo"
		End If			
		
		Set filesys = CreateObject("Scripting.FileSystemObject")
		
		filesys.CreateFolder(EXEC.Item("sExecutionResultFolder"))
		filesys.CreateFolder(EXEC.Item("sLogFolder"))
		filesys.CopyFolder cTESTREPORTLOGO , EXEC.Item("sLogFolder")

		sExecutionHtmlPath = EXEC.Item("sExecutionResultFolder") & "\ExecutionResult.html"
		Exec.Item("ExecutionResultHtml") = sExecutionHtmlPath 
		Environment.Value("ExecutionResultHtml") = Exec.Item("ExecutionResultHtml")
		
		Set filetxt = filesys.createTextFile(sExecutionHtmlPath , true)
		
		logoPath=cROOTPATH &"Layer_Test\Logo\logo.png"
		'strFileHeading="<span style='position:absolute;z-index:3;left:36px;top:1px;width:1 px;height:76px'><table cellpadding=0 cellspacing=0><tr><td width=651 height=76 valign=middle style='vertical-align:top'><![endif]>	  <div v:shape=""_x0000_s1028"" style='padding:0pt 0pt 0pt 0pt' class=shape>  <p class=Heading010 style='margin-bottom:0pt'><span lang=en-US  style='font-size:36.0pt;font-family:""Goudy Old Style"";font-weight:bold;	  language:en-US'><span style='position:absolute;left:5px;top:50px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:60.0pt;font-family:Arial;color:#888888;font-weight:normal;language:en-US'>"+projectName+"</span></p></td></tr></table></span></span></p>	  </div> <![if !vml]></td> </tr></table></span>"
		compay_logo="<span style='position:absolute;left:1200px;top:100px'><img src="""+logoPath+"""></span>"
	
		strFileHeading= "<span style='position:absolute;left:25px;top:85px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:50pt;font-family:Impact;color:#854442;font-weight:Normal;language:en-US'>"+projectName+"</span></p></td></tr></table></span>"
		
		strFileHeading2= "<span style='position:absolute;left:30px;top:160px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>"+projectTagLine+"</span></p></td></tr></table></span>"
		
		strBox= "<span style='position:absolute;left:0px;top:185px'><table border=0 style='border-collapse:collapse'><tr><td width= 2000 height=25  bgcolor= ""#4b3832""></td></tr></table></span>"
		strExecutionDate= "<span style='position:absolute;left:800px;top:190px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:13pt;font-family:Calibri;color:#fff4e6;font-weight:normal;language:en-US'>EXECUTION DATE:</span></p></td></tr></table></span>"
		
		
		strStart= "<span style='position:absolute;left:30px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>START DATE & TIME:</span></p></td></tr></table></span>"
		strEnd= "<span style='position:absolute;left:400px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>END DATE & TIME:</span></p></td></tr></table></span>"
		strTotal= "<span style='position:absolute;left:750px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>TOTAL TIME:</span></p></td></tr></table></span>"
		eStart=formatDateTime(Now)
		strStartValue= "<span style='position:absolute;left:170px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>" +cstr(eStart)+"</span></p></td></tr></table></span>"
		strEndValue= "<span style='position:absolute;left:540px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>eEnd</span></p></td></tr></table></span>"
		strTotalValue= "<span style='position:absolute;left:840px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>etotalTime</span></p></td></tr></table></span>"
		
		strTotalTest= "<span style='position:absolute;left:950px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>TOTAL TESTS:</span></p></td></tr></table></span>"
		strTotalTestValue= "<span style='position:absolute;left:1045px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>totalTX</span></p></td></tr></table></span>"
		
		strTotalPassed= "<span style='position:absolute;left:1110px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>PASSED:</span></p></td></tr></table></span>"
		strTotalPassedValue= "<span style='position:absolute;left:1170px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>totalPX</span></p></td></tr></table></span>"
		
		
		
		strTotalFailed= "<span style='position:absolute;left:1220px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>FAILED:</span></p></td></tr></table></span>"
		strTotalFailedValue= "<span style='position:absolute;left:1280px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>totalFX</span></p></td></tr></table></span>"
		
		
		
		strBox2= "<span style='position:absolute;left:0px;top:240px'><table border=0 style='border-collapse:collapse'><tr><td width= 2000 height=1  bgcolor= ""#be9b7b""></td></tr></table></span>"
	
	
		'strFileHeading = "<h2 align = 'centre'> Entra Card Regression Suite Execution Report </h2>"		
'		strExecutionDate = "<h2 align = 'centre'> Date of execution : " & Date & "</h2>"
'		strExecutionStartTime = "<h2 align = 'centre'> Execution start Time : " & formatDateTime(Time,3) & "</h2>"
'		strExecutionEndTime = "<h2 align = 'centre'> Execution End Time : overallExecutionTime" & "</h2>"
		strTableHeading = "<span style='position:absolute;left:30px;top:300px'><table align = 'centre' border='1' width='' style='border-collapse:collapse ' bordercolor = ""#be9b7b"" cellpadding ='5'><tr><th bgcolor=""#be9b7b"" width='200' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>MODULE NAME</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>MODULE DESCRIPTION</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>MODULE HEALTH</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TESTS IN MODULE</th></tr>"                                                                  
		filetxt.WriteLine(compay_logo)
		filetxt.WriteLine(strFileHeading)
		filetxt.WriteLine(strFileHeading2)
		filetxt.WriteLine(strBox)
		filetxt.WriteLine(strStart)
		filetxt.WriteLine(strEnd)
		filetxt.WriteLine(strTotal)
		filetxt.WriteLine(strStartValue)
		filetxt.WriteLine(strEndValue)
		filetxt.WriteLine(strTotalValue)
		
		filetxt.WriteLine(strTotalTest)
		filetxt.WriteLine(strTotalPassed)
		filetxt.WriteLine(strTotalFailed)
		
		filetxt.WriteLine(strTotalTestValue)
		filetxt.WriteLine(strTotalPassedValue)
		filetxt.WriteLine(strTotalFailedValue)
		
		filetxt.WriteLine(strBox2)	
		filetxt.WriteLine(strTableHeading)
		
		filetxt.Close
		
		Set filesys = Nothing
		Set filetxt = Nothing				

End sub

''' Ratnesh - End

Sub subGenerateResultLog(EXEC, sExecController, sTaskName )

	On Error Resume Next
	Dim sRootFolder,  objExcel , vFileSysObj, objWorkBook, TaskResultFolder, objWorkSheet
	  
	sRootFolder = cROOTPATH

	TimeStamp = replace (date,"/","-")&"_"& Hour(Now)&"."&Minute(Now)&"."&Second(Now)
	EXEC.Item("TASKRESULTFOLDER") = sRootFolder & "TestResultLogs\" & sTaskName & "_"& TimeStamp 
	EXEC.Item("TASKRESULTFILE") = sTaskName & "_ResultLog_"&TimeStamp&".xls"
    
	Set objExcel = CreateObject("EXCEL.APPLICATION")
	Set vFileSysObj= CreateObject("Scripting.FileSystemObject")
	Set TaskResultFolder= vFileSysObj.CreateFolder(EXEC.Item("TASKRESULTFOLDER"))

	Set objWorkBook = objExcel.Workbooks.Open(sRootFolder &"\Layer_Test\TestConfig\" & sExecController)' & ".xls")

	objWorkBook.SaveAs (TaskResultFolder&"\" & EXEC.Item("TASKRESULTFILE"))'sTaskName & "_ResultLog_"&TimeStamp&".xls")

	Set objWorkSheet= objWorkBook.Sheets("tbl_testcase")'(vTaskFolder)

	Col = 1
	While (objWorkSheet.Cells(1, Col).value <> "")
		Col =Col +1
	Wend

	objWorkSheet.Cells(1, Col).value = "Status"
	objWorkSheet.Cells(1, Col+1).value = "TimeStamp"
	objWorkSheet.Cells(1, Col+2).value = "Remark"
	objWorkSheet.Cells(1, Col+3).value = "TestFlow"
	objWorkSheet.Cells(1, Col+4).value = "SnapShots"

	
	 objWorkSheet.Range("A1:H1").Interior.ColorIndex = 15
	 objWorkSheet.Range("A1:H1").Font.Bold=15

	objWorkSheet.Cells.Select
    objWorkSheet.Cells.EntireColumn.AutoFit
    objWorkSheet.Cells.EntireColumn.AutoFit
	objWorkSheet.Columns("E:E").ColumnWidth = 15
    objWorkSheet.Columns("F:F").ColumnWidth =  65 
    objWorkSheet.Columns("G:G").ColumnWidth = 20
    objWorkSheet.Columns("H:H").ColumnWidth = 25

	objWorkBook.Save
	objWorkBook.Close True
	Set objWorkBook = Nothing
	Set objExcel = Nothing

End Sub


'Function to Format the Result Sheet generated : 
'Input paramter : Path of Result Sheet 
Sub ReusltSheetFormatting( EXEC )

	'Declaration.
		Dim objExcel ,vFileSysObj,objWorkBook,objWorkSheet,Col,Row


		
		vResultSheetpath =EXEC.Item("TASKRESULTFOLDER") &"\"& EXEC.Item("TASKRESULTFILE") 'Please Specify the result sheet path
	
	'Object Creation.
		Set objExcel = CreateObject("EXCEL.APPLICATION")
		
		Set objWorkBook = objExcel.Workbooks.Open(vResultSheetpath)
	
		Set objWorkSheet= objWorkBook.Sheets("tbl_testcase")


	'To Search "SnapShots" Column No.

		Col = 1
	
		While (objWorkSheet.Cells(1, Col).value <> "SnapShots")
			Col =Col +1
		Wend


	'To Set Hyperlink for all the Row
		Row = 2
		While (objWorkSheet.Cells(Row, Col).value <> "")
	
			objWorkBook.Sheets("tbl_testcase").Select
			objWorkSheet.Cells(Row, Col).Select
			vValue = objWorkSheet.Cells(Row, Col).value 
			objWorkBook.ActiveSheet.Hyperlinks.Add objWorkSheet.Cells(Row, Col),vValue ' Code to Set Hyperlink
			Row = Row +1
	
		Wend

	'To select all Cells and set Text Wrap as ON

		objWorkSheet.Cells.WrapText = True

	
	'Save and Close Objects
		objWorkBook.Save
		objWorkBook.Close True
		Set objWorkBook = Nothing
		Set objExcel = Nothing



End Sub


Function subGenerateResultLog_Swap_OLD(EXEC, sExecController, sTaskName )

    On Error Resume Next
    Dim sRootFolder,  objExcel , vFileSysObj, objWorkBook, TaskResultFolder, objWorkSheet
      
    sRootFolder = cROOTPATH

    TimeStamp = replace (date,"/","-")&"_"& Hour(Now)&"."&Minute(Now)&"."&Second(Now)
    'EXEC.Item("TASKRESULTFOLDER") = sRootFolder & "TestResultLogs\" & sTaskName & "_" & TimeStamp ' Module Result Folder
    'EXEC.Item("TASKRESULTFILE") = sTaskName & "_ResultLog_"&TimeStamp&".xls"    ' Inner Result file

    Dim sResultFolderPath,strTime,filesys,sExecutionFolder,sExecutionFolderPath,strTaskResultFolder,strTableHeading,strFileHeading,strExecutionDate,strExecutionTime

    strTime = Hour(Now) & "::" & Minute(Now) & "::" & Second(Now)
    sTaskName = Replace(sTaskName , " " , "_")

    ' Check and Create the execution Folder

    Set filesys = CreateObject("Scripting.FileSystemObject")
    
    sExecutionFolder = "_ResultLog_" '& TimeStamp
    sExecutionFolderPath = sRootFolder & "TestResultLogs\" & sTaskName & "_" & TimeStamp & "\" & sExecutionFolder 

    
  	' Exec.Item("ModuleResultFolder") = sExecutionFolderPath 
 	sExecutionFolderPath= EXEC.Item("TCFOLDER")


    if filesys.FolderExist(sExecutionFolderPath ) = false Then
          filesys.createFolder(sExecutionFolderPath )
    End If

    sResultHtmlFilePath = sExecutionFolderPath  & "\" & EXEC.Item("TCId") &"HTML Result" & ".html"
    Exec.Item("ModuleResultHtml") = sResultHtmlFilePath 
	Environment.Value("ModuleResultHtml") = Exec.Item("ModuleResultHtml")

    Set filetxt = filesys.createTextFile(sResultHtmlFilePath , true)

	









'    strFileHeading = "<h2 align = 'centre'> Module : " & Replace(EXEC.Item("TCId"), "_" , " " ) & "</h2>"     
'    strExecutionDate = "<h2 align = 'centre'> Date of execution : " & Date & "</h2>"
'    strExecutionStartTime = "<h2 align = 'centre'> Execution start Time : " & formatDateTime(Time,3) & "</h2>"
'    strExecutionEndTime = "<h2 align = 'centre'> Execution End Time : moduleExecutionTime" & "</h2>"
'    strTableHeading = "<table align = 'centre' border='1' width='' bordercolor = 'black'><tr><th bgcolor='silver' width='50'>Step No.</th><th bgcolor='silver'  width='300'>Step Description </th><th bgcolor='silver' width='300'>Actual</th><th bgcolor='silver'>Status</th><th bgcolor='silver'>Snapshot</th></tr> </table>"                                                                  
'    
'    filetxt.WriteLine(strFileHeading) 
'    filetxt.WriteLine(strExecutionDate)
'    filetxt.WriteLine(strExecutionStartTime)
'    filetxt.WriteLine(strExecutionEndTime)
'    filetxt.WriteLine(strTableHeading)

    Set filesys = Nothing
    Set filetxt = Nothing  

    subGenerateResultLog_Swap = sResultHtmlFilePath 

End Function


Function subGenerateResultLog_Swap(EXEC, sExecController, sTaskName )

    On Error Resume Next
    Dim sRootFolder,  objExcel , vFileSysObj, objWorkBook, TaskResultFolder, objWorkSheet
      
    sRootFolder = cROOTPATH

    'TimeStamp = replace (date,"/","-")&"_"& Hour(Now)&"."&Minute(Now)&"."&Second(Now)
    'EXEC.Item("TASKRESULTFOLDER") = sRootFolder & "TestResultLogs\" & sTaskName & "_" & TimeStamp ' Module Result Folder
    'EXEC.Item("TASKRESULTFILE") = sTaskName & "_ResultLog_"&TimeStamp&".xls"    ' Inner Result file

    Dim sResultFolderPath,strTime,filesys,sExecutionFolder,sExecutionFolderPath,strTaskResultFolder,strTableHeading,strFileHeading,strExecutionDate,strExecutionTime

	sResultFolderPath = EXEC.Item("sExecutionResultFolder")

    strTime = Hour(Now) & "::" & Minute(Now) & "::" & Second(Now)
    sTaskName = Replace(sTaskName , " " , "_")

    ' Check and Create the execution Folder

    Set filesys = CreateObject("Scripting.FileSystemObject")
    
    sExecutionFolder = "Results_" & sTaskName & "_" & Day(date) & "_" & Month(date) & "_" & Year(date) & "_" & Replace(strTime , "::" , "_")
    sExecutionFolderPath = sResultFolderPath & "\" & sExecutionFolder 

	EXEC.Item("ModuleResultFolder") = sExecutionFolderPath    

    if filesys.FolderExists(sExecutionFolderPath ) = false Then
          filesys.createFolder(sExecutionFolderPath )
    End If

    sResultHtmlFilePath = sExecutionFolderPath  & "\" & sTaskName & ".html"
    Exec.Item("ModuleResultHtml") = sResultHtmlFilePath 

    Set filetxt = filesys.createTextFile(sResultHtmlFilePath , true)
    
    '	logoPath=cROOTPATH &"Layer_Test\Logo\logo.png"
		strFileHeading="<span style='position:absolute;z-index:3;left:36px;top:1px;width:1 px;height:76px'><table cellpadding=0 cellspacing=0><tr><td width=651 height=76 valign=middle style='vertical-align:top'><![endif]>	  <div v:shape=""_x0000_s1028"" style='padding:0pt 0pt 0pt 0pt' class=shape>  <p class=Heading010 style='margin-bottom:0pt'><span lang=en-US  style='font-size:36.0pt;font-family:""Goudy Old Style"";font-weight:bold;	  language:en-US'><span style='position:absolute;left:5px;top:50px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:40.0pt;font-family:Arial;color:#888888;font-weight:normal;language:en-US'>EBS AUTOMATION </span></p></td></tr></table></span></span></p>	  </div> <![if !vml]></td> </tr></table></span>"
	'	compay_logo="<span style='position:absolute;left:1000px;top:30px'><img src="""+logoPath+"""></span>"
	
		strFileHeading= "<span style='position:absolute;left:25px;top:100px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:40pt;font-family:Impact;color:#854442;font-weight:Normal;language:en-US'>Module : " + Ucase (Replace(sTaskName , "_" , " " ))+"</span></p></td></tr></table></span>"
		
		strFileHeading2= "<span style='position:absolute;left:30px;top:160px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>THE DETAILS OF THE TESTCASES IN MODULE '" + Ucase (Replace(sTaskName , "_" , " " ))+"' CAN BE SEEN BELOW</span></p></td></tr></table></span>"
		
		strBox= "<span style='position:absolute;left:0px;top:185px'><table border=0 style='border-collapse:collapse'><tr><td width= 2000 height=25  bgcolor= ""#4b3832""></td></tr></table></span>"
		strExecutionDate= "<span style='position:absolute;left:800px;top:190px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:13pt;font-family:Calibri;color:#fff4e6;font-weight:normal;language:en-US'>EXECUTION DATE:</span></p></td></tr></table></span>"
		
		
		strStart= "<span style='position:absolute;left:30px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>START DATE & TIME:</span></p></td></tr></table></span>"
		strEnd= "<span style='position:absolute;left:400px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>END DATE & TIME:</span></p></td></tr></table></span>"
		strTotal= "<span style='position:absolute;left:750px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#3c2f2f;font-weight:normal;language:en-US'>TOTAL TIME:</span></p></td></tr></table></span>"
		mStart=formatDateTime(Now)
		strStartValue= "<span style='position:absolute;left:170px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>" +cstr(mStart)+"</span></p></td></tr></table></span>"
		strEndValue= "<span style='position:absolute;left:540px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>mEnd</span></p></td></tr></table></span>"
		strTotalValue= "<span style='position:absolute;left:840px;top:212px'><table><tr><td><p class=MsoNormal><span lang=en-US style='font-size:12pt;font-family:Calibri;color:#be9b7b;font-weight:normal;language:en-US'>mtotalTime</span></p></td></tr></table></span>"
		strBox2= "<span style='position:absolute;left:0px;top:240px'><table border=0 style='border-collapse:collapse'><tr><td width= 2000 height=1  bgcolor= ""#be9b7b""></td></tr></table></span>"
	
	
		'strFileHeading = "<h2 align = 'centre'> Entra Card Regression Suite Execution Report </h2>"		
'		strExecutionDate = "<h2 align = 'centre'> Date of execution : " & Date & "</h2>"
'		strExecutionStartTime = "<h2 align = 'centre'> Execution start Time : " & formatDateTime(Time,3) & "</h2>"
'		strExecutionEndTime = "<h2 align = 'centre'> Execution End Time : overallExecutionTime" & "</h2>"
'		strTableHeading = "<span style='position:absolute;left:30px;top:300px'><table align = 'centre' border='1' width='' style='border-collapse:collapse ' bordercolor = ""#be9b7b"" cellpadding ='5'><tr><th bgcolor=""#be9b7b"" width='40' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TC.NO.</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TESTCASE NAME</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TESTCASE DESCRIPTION</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TESTCASE STATUS</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TEST WORKFLOW</th></tr>"                                                                  
		strTableHeading = "<span style='position:absolute;left:30px;top:300px'><table align = 'centre' border='1' width='' style='border-collapse:collapse ' bordercolor = ""#be9b7b"" cellpadding ='5'><tr><th bgcolor=""#be9b7b"" width='40' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TC.NO.</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TESTCASE NAME</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TESTCASE DESCRIPTION</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>BU</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TESTCASE STATUS</th><th bgcolor=""#be9b7b""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>TEST WORKFLOW</th></tr>"
	
	'	filetxt.WriteLine(compay_logo)
		filetxt.WriteLine(strFileHeading)
		filetxt.WriteLine(strFileHeading2)
		filetxt.WriteLine(strBox)
		filetxt.WriteLine(strStart)
		filetxt.WriteLine(strEnd)
		filetxt.WriteLine(strTotal)
		filetxt.WriteLine(strStartValue)
		filetxt.WriteLine(strEndValue)
		filetxt.WriteLine(strTotalValue)
		filetxt.WriteLine(strBox2)	
		filetxt.WriteLine(strTableHeading)
		
		filetxt.Close

'
'    strFileHeading = "<h2 align = 'centre'> Module : " & Replace(sTaskName , "_" , " " ) & "</h2>"     
'    strExecutionDate = "<h2 align = 'centre'> Date of execution : " & Date & "</h2>"
'    strExecutionStartTime = "<h2 align = 'centre'> Execution start Time : " & formatDateTime(Time,3) & "</h2>"
'    strExecutionEndTime = "<h2 align = 'centre'> Execution End Time : moduleExecutionTime" & "</h2>"
'    strTableHeading = "<table align = 'centre' border='1' width='' bordercolor = 'black'><tr><th bgcolor='silver' width='30'>TC No.</th><th bgcolor='silver'  width='300'>Test Case Name</th><th bgcolor='silver' width='300'>Test Case Description</th><th bgcolor='silver'>TC Status</th><th bgcolor='silver'>Analyse Test</th></tr>"                                                                  
'    
'    filetxt.WriteLine(strFileHeading) 
'    filetxt.WriteLine(strExecutionDate)
'    filetxt.WriteLine(strExecutionStartTime)
'    filetxt.WriteLine(strExecutionEndTime)
'    filetxt.WriteLine(strTableHeading)
'	
'	filetxt.close
	
    Set filesys = Nothing
    Set filetxt = Nothing  

    subGenerateResultLog_Swap = sResultHtmlFilePath 

End Function




Sub subWriteTestStepResult(takeSnap,EXEC, iStepCount , sStepDescription , sPassActual, sFailActual)

    On Error Resume Next

    Dim objFsys , objFiletxt, sImagePath , strTestStepResultLine , sImageRelativePath
	Dim linkTag
    
    If IsNumeric ( iStepCount ) Then               ' Fail Step
       iStepCount = iStepCount + 1
    End If

    If iStepCount = "" or bResultIndicator =  True Then
    
        Set objFsys = CreateObject("Scripting.FileSystemObject")
        Set objFiletxt = objFsys.OpenTextFile(EXEC.Item("TestCaseResultHtmlPath") , 8 , True )
		
		If lcase (takeSnaps)="all" Then
			sImagePath = EXEC.Item("sSnapshotFolder") & "\Snapshot" & iStepCount & ".png"        
        	Desktop.CaptureBitMap(sImagePath)        
        	sImageRelativePath = "..\" & EXEC.Item("TCFOLDERName") & Split( sImagePath , EXEC.Item("TCFOLDERName") ) (1)
        	linkTag="<a href=""" & sImageRelativePath & """>Snapshot..."			
		Else
			If bKeywordPF=false Then
				sImagePath = EXEC.Item("sSnapshotFolder") & "\Snapshot" & iStepCount & ".png"        
	        	Desktop.CaptureBitMap(sImagePath)        
	        	sImageRelativePath = "..\" & EXEC.Item("TCFOLDERName") & Split( sImagePath , EXEC.Item("TCFOLDERName") ) (1)
	        	linkTag="<a href=""" & sImageRelativePath & """>Snapshot..."
			ElseIf takeSnap=true Then
				sImagePath = EXEC.Item("sSnapshotFolder") & "\Snapshot" & iStepCount & ".png"        
	        	Desktop.CaptureBitMap(sImagePath)        
	        	sImageRelativePath = "..\" & EXEC.Item("TCFOLDERName") & Split( sImagePath , EXEC.Item("TCFOLDERName") ) (1)
	        	linkTag="<a href=""" & sImageRelativePath & """>Snapshot..."			
			Else
				linkTag="Not taken..."		
				
			End If
				
		End If

        


		If iStepCount mod 2 <> 0 Then
			tempBGColor="#fff4e6"
		else
			tempBGColor="#fff4e6"
		End If

        if bKeywordPF = True Then
        	strTestStepResultLine = "<tr><td  width='40'align=""center"" bgcolor='"+tempBGColor+"' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & iStepCount & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sStepDescription & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sPassActual & "</td><td bgcolor='Green' align=""center""  style='font-size:12pt;font-family:Calibri; color:#000000;font-weight:normal;language:en-US'><span style='color:#ffffff'>PASS</td></span></td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>"+linkTag+"</td></tr>"
'        	iPassCount = iPassCount + 1
        Else        
            strTestStepResultLine = "<tr><td  width='40'align=""center"" bgcolor='"+tempBGColor+"'style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & iStepCount & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sStepDescription & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sFailActual & "</td><td bgcolor='Red' align=""center""  style='font-size:12pt;font-family:Calibri; color:#000000;font-weight:normal;language:en-US'><span style='color:#ffffff'>FAIL</td></span></td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>"+linkTag+"</td></tr>"
            bResultIndicator = false
'			iFailCount = iFailCount + 1            
        End If


        objFiletxt.WriteLine(strTestStepResultLine)
        objFiletxt.close
        Set objFiletxt = nothing

    End if
    
End Sub


'Sub subWriteTestCaseResult(EXEC , TCCount , TCNAME , TCDescription)
'	
'	Dim objFsys , objFiletxt, strTestResultLine , sTCRelativeFilePath	
'	
'	Set objFsys = CreateObject("Scripting.FileSystemObject")	
'	Set objFiletxt = objFsys.OpenTextFile(Exec.Item("ModuleResultHtml") , 8 , True )
'
'	sTCRelativeFilePath = ".." & Split( Exec.Item("TestCaseResultHtmlPath") , Exec.Item("sExecutionResultFolderName") )(1)	
'	
'	If TCCount mod 2 = 0 Then
'		tempBGColor="#fff4e6"
'	else
'		tempBGColor="#fff4e6"
'	End If
'	
'	If bMethodPF = True And bKeywordPF = True Then
'
'		strTestResultLine = "<tr><td width='40' align=""center"" bgcolor='"+tempBGColor+"' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCCount+1 & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCNAME & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCDescription & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sDataSetColumn & "</td><td bgcolor='Green' align=""center""  style='font-size:12pt;font-family:Calibri; color:#000000;font-weight:normal;language:en-US'><span style='color:#ffffff'>PASS</td></span><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><a href=" & sTCRelativeFilePath & ">Steps in test...</td></tr>"    
'		iPassCount = iPassCount + 1
'		iPassCountSuite=iPassCountSuite+1
'	Else
'		strTestResultLine = "<tr><td width='40' align=""center"" bgcolor='"+tempBGColor+"' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCCount+1 & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCNAME & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCDescription & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sDataSetColumn & "</td><td bgcolor='Red' align=""center""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><span style='color:#ffffff'>FAIL</td></span></td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><a href=" & sTCRelativeFilePath & ">Steps in test...</td></tr>"   	                                                           
'		iFailCount = iFailCount + 1 
'		iFailCountSuite=iFailCountSuite+1		
'	End If
'	
'	objFiletxt.WriteLine(strTestResultLine)
'	objFiletxt.Close
'	Set objFiletxt = nothing
'	Set objFsys = nothing 
'	
'End Sub


Sub subWriteModuleResult(EXEC,ModuleName,ModuleDescription)
	
	Dim objFsys , objFiletxt, sTCRelativeFilePath , iTotalTC 
	Dim iPassCounter , strTestResultLineStart , strTestResultLineEnd , strTestResultLineHealth , iHealthCounter

	Set objFsys = CreateObject("Scripting.FileSystemObject")	
	Set objFiletxt = objFsys.OpenTextFile(Exec.Item("ExecutionResultHtml") , 8 , True )

	sTCRelativeFilePath = ".." & Split( Exec.Item("ModuleResultHtml") , "TestResultLogs" )(1)
	iTotalTC = iPassCount + iFailCount
	iPassCounter = Round( ( (iPassCount/iTotalTC) * 10) , 0 )

	strTestResultLineStart = "<tr><td width='200' bgcolor='#fff4e6' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US' >" & ModuleName & "</td><td bgcolor='#fff4e6' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & ModuleDescription & "</td><td bgcolor='#fff4e6' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><table border='0' width='100%' bordercolor='' align='center' cellspacing='0'></tr>"
	objFiletxt.WriteLine(strTestResultLineStart)

	For iHealthCounter = 1 to iPassCounter
		objFiletxt.WriteLine("<td bgcolor='green' height='10' width='10'></td>")	
	Next

	For iHealthCounter = 1 to 10-iPassCounter
		objFiletxt.WriteLine("<td bgcolor='red' height='10' width='10'></td>")		
	Next

	strTestResultLineEnd = "</tr></table></td><td bgcolor='#fff4e6' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><a href=" & sTCRelativeFilePath & ">Click to see tests...</td></tr>"
	objFiletxt.WriteLine(strTestResultLineEnd)

	objFiletxt.Close
	Set objFiletxt = nothing
	Set objFsys = nothing

End Sub




Public Function getOBJ(RepositoryFrom,allObjectsCollection,ParentObject)

'Get Objects by parent From loaded Repository
'If parent not specified all objects will be returned

Set fTOCollection = RepositoryFrom.GetChildren(ParentObject)

    For RepObjIndex = 0 To fTOCollection.Count - 1

        'Get object by index
        Set fTestObject = fTOCollection.Item(RepObjIndex)

            'Check whether the object is having child objects
            If RepositoryFrom.GetChildren (fTestObject).count<>0 then

				ObjectName = RepositoryFrom.GetLogicalName(fTestObject) ' Object's Logical name
				ObjectClass = fTestObject.GetTOProperty("micclass") 
				ObjectDefination=ObjectClass&"("""&ObjectName&""")"
				
					Dim pppath
					pppath=getPath(RepositoryFrom,allObjectsCollection,fTestObject,"","","No",ObjectDefination)
					print now
					print Timer()

					print pppath
				
				OBJ.Add ObjectName,pppath ' Add the parent of the Node 
				'print OBJ.item(ObjectName)
						If Err.Number <> 0 Then
							msgbox "OR issue--" & sNodeLabel & "-->" & Err.description
							Reporter.ReportEvent micFail,"OR issue",sNodeLabel & "-->" & Err.description
							ExitTest				
						End If


                getOBJ RepositoryFrom,allObjectsCollection,fTestObject',path',2

            else
                
                ObjectName = RepositoryFrom.GetLogicalName(fTestObject) ' Object's Logical name
				ObjectClass = fTestObject.GetTOProperty("micclass") 
				ObjectDefination=ObjectClass&"("""&ObjectName&""")"
					Dim pppath1
					pppath1=getPath(RepositoryFrom,allObjectsCollection,fTestObject,"","","No",ObjectDefination)
					print pppath1
					print now
					print Timer()
					OBJ.Add ObjectName,pppath1 ' Add the parent of the Node 
						If Err.Number <> 0 Then
							msgbox "OR issue --" & sNodeLabel & "-->" & Err.description
							Reporter.ReportEvent micFail,"OR issue",sNodeLabel & "-->" & Err.description
							ExitTest				
						End If					
            End if
            
           
    Next

End Function





'Function getPath(RepositoryFrom,fTestObject,topMostParent,treePath,objFound,ObjectDefination)
'
'
'	Dim tempObject	
'		Set allObjectsCollection = RepositoryFrom.GetAllObjects
'		totalObjects=allObjectsCollection.Count
'		ub1=totalObjects - 1
'		 For p = 0 To ub1
'		 	 Set tempObject = allObjectsCollection.Item(p)
'		 	 		Set allChildrenCollection = RepositoryFrom.GetChildren(tempObject)
'		 	 		childCount=allChildrenCollection.Count
'		 			 if (childCount<>0) then
'			 			 'Set allChildrenCollection = RepositoryFrom.GetChildren(tempObject)
'			 			 ub2=childCount-1
'			 			 For t = 0 To childCount
'			 			 	 Set childName = allChildrenCollection.Item(t)
'			 			 	 If RepositoryFrom.GetLogicalName(childName)=RepositoryFrom.GetLogicalName(fTestObject) Then
'			 			   		objFound="yes"
'			 			  		Exit for
'			 			  	End If
'			 			 Next 			 
'		 			 End if 
'		 	 If objFound="yes" Then 
'		 	 	Exit for
'		 	 End If
' 		Next
' 		
' if(objFound="yes") then	
'		tempObjectName = RepositoryFrom.GetLogicalName(tempObject) ' Object's Logical name
'		tempObjectClass = tempObject.GetTOProperty("micclass") 
'		treePath= tempObjectClass&"("""&tempObjectName&""")."+treePath	
'		getPath RepositoryFrom,tempObject,topMostParent,treePath,"No",ObjectDefination
'	
'  End If
' 
' getPath=treePath+ObjectDefination
' 
'End Function
'



'=================== MOMIN WASE

Function getPath( RepositoryFrom,  allObjectsCollection, fTestObject,topMostParent,treePath,objFound, ObjectDefination)

	Dim tempObject	
		'Set allObjectsCollection = allObjects 'assigned to global variable to reduce time 
		totalObjects=allObjectsCollection.Count
		ub1=totalObjects - 1
		 For p = 0 To ub1
		 	 Set tempObject = allObjectsCollection.Item(p)
		 	 		'print RepositoryFrom.GetLogicalName(tempObject)
		 			 Set allChildrenCollection = RepositoryFrom.GetChildren(tempObject)
		 			 childCount=allChildrenCollection.Count
		 			 if (childCount<>0) then			 			 			
			 			 For t = 0 To childCount- 1
			 			 	 Set childName = allChildrenCollection.Item(t)
			 			 	 If RepositoryFrom.GetLogicalName(childName)=RepositoryFrom.GetLogicalName(fTestObject) Then
			 			   		objFound="yes"
			 			  		Exit for
			 			  	End If
			 			  	Set childName =nothing
			 			 Next 
					 	 Set allChildrenCollection = nothing				 			 
		 			 End if 
		 	 If objFound="yes" Then 
		 	 	Exit for
		 	 End If
 		Next
 		
 if(objFound="yes") then	
		tempObjectName = RepositoryFrom.GetLogicalName(tempObject) ' Object's Logical name
		tempObjectClass = tempObject.GetTOProperty("micclass") 
		treePath= tempObjectClass&"("""&tempObjectName&""")."+treePath	
		getPath RepositoryFrom,allObjectsCollection,tempObject,topMostParent,treePath,"No",ObjectDefination		
		Set tempObject =nothing
	
  End If
 
 getPath=treePath+ObjectDefination
 
End Function

'
'Function getPath(RepositoryFrom,fTestObject,topMostParent,treePath,objFound,ObjectDefination)
'
'
'	Dim tempObject	
'		Set allObjectsCollection = RepositoryFrom.GetAllObjects
'		 For p = 0 To allObjectsCollection.Count - 1
'		 	 Set tempObject = allObjectsCollection.Item(p)
'		 			 if (RepositoryFrom.GetChildren(tempObject).count<>0) then
'			 			 Set allChildrenCollection = RepositoryFrom.GetChildren(tempObject)
'			 			 For t = 0 To allChildrenCollection.Count - 1
'			 			 	 Set childName = allChildrenCollection.Item(t)
'			 			 	 If RepositoryFrom.GetLogicalName(childName)=RepositoryFrom.GetLogicalName(fTestObject) Then
'			 			   		objFound="yes"
'			 			  		Exit for
'			 			  	End If
'			 			 Next 			 
'		 			 End if 
'		 	 If objFound="yes" Then 
'		 	 	Exit for
'		 	 End If
' 		Next
' 		
' if(objFound="yes") then	
'		tempObjectName = RepositoryFrom.GetLogicalName(tempObject) ' Object's Logical name
'		tempObjectClass = tempObject.GetTOProperty("micclass") 
'		treePath= tempObjectClass&"("""&tempObjectName&""")."+treePath	
'		getPath RepositoryFrom,tempObject,topMostParent,treePath,"No",ObjectDefination
'	
'  End If
' 
' getPath=treePath+ObjectDefination
' 
'End Function



	
	
'	Function getOBJFromFile()		
'	
'	'filename = "C:\Temp\vblist.txt"
'
'Set fso = CreateObject("Scripting.FileSystemObject")
''Set f = fso.OpenTextFile(filename)
''
''Do Until f.AtEndOfStream
''  WScript.Echo f.ReadLine
''Loop
''
''f.Close
'		
'		Set listFile = fso.OpenTextFile("C:\AUTOMATION_ATLAS\Layer_Application\OR\OR_Description.qfl")  ' <-- remove .ReadAll from this line!
'		Do Until listFile.AtEndOfStream
'		  fName = listFile.ReadLine
'		  print fName
'		Loop
'	End Function



'**************************************************************************************************************************************************************************************
'CODE FOR NEW EXECUTION CONTROL
'**************************************************************************************************************************************************************************************


Function kyGetAPITestDetails(EXEC, UNIVERSAL)
	Dim arrAPITestDetails(4)
	'strVariableName = EXEC.Item("TCId") & "_" & "APITestNumber"
	varAPIActionName = UNIVERSAL.item("varAPIActionName")
	varAPIScriptName = UNIVERSAL.item("APIScriptName")
	arrAPITestDetails(0) = varAPIActionName
	arrAPITestDetails(1) = varAPIScriptName
	kyGetAPITestDetails = arrAPITestDetails
End Function

Function InitializeSuitDetails()		
        Set UNIVERSAL = CreateObject ("Scripting.Dictionary")
        Set EXEC = CreateObject("Scripting.Dictionary")	    
        EXEC.RemoveAll
        UNIVERSAL.RemoveAll		           
        subCreateExecutionResultFolder cTESTRESULTLOGFOLDERPATH, EXEC '''Creating the top level html structure
        sExecController = mController
        sEnv ="TestSet1"        

	InitializeSuitDetails = sExecController
End Function

Function EndTheSuit()
Dim oFSO, oF, fileContent

	 Set oFSO  = createobject("scripting.Filesystemobject")
        
        If oFSO.FileExists(Exec.Item("ExecutionResultHtml")) Then
            Set oF = oFSO.OpenTextFile(Exec.Item("ExecutionResultHtml"))
            fileContent = oF.ReadAll
            eEnd=formatDateTime(Now)
            fileContent = Replace(fileContent , "eEnd" , cstr(eEnd) )
            fileContent = Replace(fileContent , "etotalTime" ,  SplitSec(DateDiff("s",eStart,eEnd)))
            
            fileContent = Replace(fileContent , "totalTX" , cstr(iPassCountSuite+iFailCountSuite))
            fileContent = Replace(fileContent , "totalPX" , cstr(iPassCountSuite))
            fileContent = Replace(fileContent , "totalFX" , cstr(iFailCountSuite))
            
            oF.Close
            Set oF = nothing            
            Set oF = oFSO.OpenTextFile(Exec.Item("ExecutionResultHtml") , 2 )
            oF.Write fileContent
            oF.Close
            Set oF = nothing            
        End If

        If oFSO.FileExists(Environment.Value("ExecutionResultHtml")) Then
            systemutil.Run Environment.Value("ExecutionResultHtml")
'            Sub_UploadResultoALM(Environment.Value("ExecutionResultHtml") ) 								'Upload Suite test result html file to the Testset in ALM
'	    	 Sub_DisconnectFromALM	
        End If

        Set oFSO = Nothing
End Function


Function IntializeTestcaseSetup(sExecController)
	 Dim arrTestCases, iTestCaseCount, arrModule,iModuleCount
    Dim qtpApp ,qtpRepositories
    Dim strName,strOrName,strOrPath , oFSO , oF , fileContent	      
    'Set qtpApp = CreateObject("QuickTest.Application")   
    Dim anyORUploaded
    anyORUploaded=false
   
    Dim ParentObject
    Dim RepositoryFrom    
	Set	OBJ=CreateObject("Scripting.Dictionary")
	OBJ.RemoveAll
	
	'Create Object Rpository and load to OBJ dictionary if the OR Type is "ProjectLevel"
	If ORType="ProjectLevel" Then	
		Dim projectORFilePath
		projectORFilePath=ORFolderPath & "\" & projectOR
		Set qtpRepositories = qtpApp.Test.Actions("Action1").ObjectRepositories 
	
	    qtpRepositories.Add(projectORFilePath)
	    anyORUploaded=true	
		Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")
		RepositoryFrom.Load projectORFilePath
		set allObjectsCollection=RepositoryFrom.GetAllObjects
		getOBJ RepositoryFrom,allObjectsCollection,ParentObject',""',1
		
	End If    
'	getOBJFromFile
    arrModule =fGetProjectModules(sExecController,"tbl_modules")  
    
    IntializeTestcaseSetup = arrModule
	
End Function

Function InitializeModuleSetup(arrModule, iModuleCount)
	 iPassCount = 0
        iFailCount = 0
        Environment.Value("executionFurther") = True
        EXEC.Item("ModuleResultHtml") = subGenerateResultLog_Swap(EXEC, arrModule(iModuleCount,1),arrModule(iModuleCount,0) )
        EXEC.Item("MODULEDATAFILE") = arrModule(iModuleCount,2)
                    
        'Create Object Rpository and load to OBJ dictionary if the OR Type is "ModuleWise"
		If ORType="ModuleWise" Then	
		    Dim orPath,sParentObject
		    Dim i,orNames,arrORNames
'			Set qtpRepositories = qtpApp.Test.Actions("Action1").ObjectRepositories 
'			Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")
			 
			orNames = arrModule(iModuleCount,4)
			If (len(orNames)>0) Then
				arrORNames=split (orNames,";")	
				
				
			
			For i = 0 to UBound(arrORNames)		
				Set qtpRepositories = qtpApp.Test.Actions("Action1").ObjectRepositories 
				Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")			
				orPath=ORFolderPath & "\" & arrORNames(i)
				qtpRepositories.Add(orPath)
				anyORUploaded=true
				RepositoryFrom.Load orPath	
				set allObjectsCollection=RepositoryFrom.GetAllObjects ' This is added to rduce to time creating OBJ.item
				getOBJ RepositoryFrom,allObjectsCollection,sParentObject',""',1
				set allObjectsCollection=nothing 'memory leaks
				Set RepositoryFrom =nothing
			Next
			
			
			End If
		
		End If  

        arrTestCases = fGetTestCases( arrModule(iModuleCount,1),arrModule(iModuleCount,0))
	InitializeModuleSetup = arrTestCases
End Function

Function EndModuleSetup(arrModule,iModuleCount)

Dim oFSO, oF, fileContent
	        Set oFSO  = createobject("scripting.Filesystemobject")
        
        If oFSO.FileExists(Exec.Item("ModuleResultHtml")) Then
            Set oF = oFSO.OpenTextFile(Exec.Item("ModuleResultHtml"))
            fileContent = oF.ReadAll
            mEnd=formatDateTime(Now)
            fileContent = Replace(fileContent , "mEnd" , cstr(mEnd) )
            fileContent = Replace(fileContent , "mtotalTime" ,  SplitSec(DateDiff("s",mStart,mEnd)) )        
      		oF.Close
            oF.Close
            Set oF = nothing
       		Set oF = oFSO.OpenTextFile(Exec.Item("ModuleResultHtml") , 2 )
            oF.Write fileContent
            oF.Close
            Set oF = nothing
         End If

        Set oFSO = Nothing    
        subWriteModuleResult EXEC , arrModule(iModuleCount,0) , arrModule(iModuleCount,3)

       	If ORType="ModuleWise" and anyORUploaded=true Then
			qtpRepositories.RemoveAll
			OBJ.RemoveAll
		End if 	
End Function


Function EndTestcaseSetup(TCNAME , TCDesription)
Dim oFSO, oF, fileContent

	Set oFSO  = createobject("scripting.Filesystemobject")

    If oFSO.FileExists(Environment.Value("TestCaseResultHtmlPath")) Then
        Set oF = oFSO.OpenTextFile(Environment.Value("TestCaseResultHtmlPath") , 1)
        fileContent = oF.ReadAll
        tEnd=formatDateTime(Now)
            fileContent = Replace(fileContent , "tEnd" , cstr(tEnd) )
            fileContent = Replace(fileContent , "ttotalTime" ,  SplitSec(DateDiff("s",tStart,tEnd)) )        
        oF.Close
        Set oF = nothing        
        Set oF = oFSO.OpenTextFile(Environment.Value("TestCaseResultHtmlPath") , 2 )
        oF.Write fileContent
        oF.Close
        Set oF = nothing        
    End If

    subWriteTestCaseResult EXEC , TCNo , TCNAME , TCDesription   
End Function


Sub subWriteTestCaseResult(EXEC , TCCount , TCNAME , TCDescription)
	
	Dim objFsys , objFiletxt, strTestResultLine , sTCRelativeFilePath	
	
	Set objFsys = CreateObject("Scripting.FileSystemObject")	
	Set objFiletxt = objFsys.OpenTextFile(Exec.Item("ModuleResultHtml") , 8 , True )

	sTCRelativeFilePath = ".." & Split( Exec.Item("TestCaseResultHtmlPath") , Exec.Item("sExecutionResultFolderName") )(1)	
	
	If TCCount mod 2 = 0 Then
		tempBGColor="#fff4e6"
	else
		tempBGColor="#fff4e6"
	End If
	
	If bMethodPF = True Then
		strTestResultLine = "<tr><td width='40' align=""center"" bgcolor='"+tempBGColor+"' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCCount+1 & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCNAME & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCDescription & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sDataSetColumn & "</td><td bgcolor='Green' align=""center""  style='font-size:12pt;font-family:Calibri; color:#000000;font-weight:normal;language:en-US'><span style='color:#ffffff'>PASS</td></span><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><a href=" & sTCRelativeFilePath & ">Steps in test...</td></tr>"   
		iPassCount = iPassCount + 1
		iPassCountSuite=iPassCountSuite+1
	Else
		strTestResultLine = "<tr><td width='40' align=""center"" bgcolor='"+tempBGColor+"' style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCCount+1 & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCNAME & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & TCDescription & "</td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'>" & sDataSetColumn & "</td><td bgcolor='Red' align=""center""  style='font-size:12pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><span style='color:#ffffff'>FAIL</td></span></td><td bgcolor='"+tempBGColor+"' style='font-size:11pt;font-family:Calibri;color:#000000;font-weight:normal;language:en-US'><a href=" & sTCRelativeFilePath & ">Steps in test...</td></tr>"   	
		iFailCount = iFailCount + 1 
		iFailCountSuite=iFailCountSuite+1		
	End If
	
	objFiletxt.WriteLine(strTestResultLine)
	objFiletxt.Close
	Set objFiletxt = nothing
	Set objFsys = nothing 
	
End Sub




Function keywordSetFetch(sDataSetName)

iDataSetNum = Split(sDataSetName, "-")(1)
sEnvironmentName = Replace(Split(sDataSetName, "-")(0),"ENV", "ENVIRONMENT")

sColumnName = Replace(sEnvironmentName, ":", "")


	keywordSetFetch = sColumnName & "_" & Cint(iDataSetNum)
End Function



Function SheetNameFetch(sTestCaseName)

sModuleName = Split(sTestCaseName, "_")(0)
sTCNum = Split(sTestCaseName, "_")(1)
	SheetNameFetch = sModuleName &"_"& sTCNum
	
End Function




