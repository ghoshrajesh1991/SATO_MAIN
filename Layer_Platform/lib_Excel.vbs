Function kyExcelDataExport(sDataBase, sDataTable, iRowId, sColumnName, sData)
on Error Resume Next
			If bKeywordPF = True Then
				strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBase & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
				sSQLQuery = "Update [" & sDataTable & "$] Set "&sColumnName &"="&sData&" where SNo = "&iRowId
							Set objConn = CreateObject("ADODB.Connection")'Create Connection Object		
							Set objRS = CreateObject("ADODB.Recordset")'Create the RecordSet	
							objConn.ConnectionString = strConnection
						    objConn.ConnectionTimeOut = 10
						    objConn.CommandTimeOut = 10
			
			    objConn.Open' Open the Connection with DB mentioned in "strConnection"
			    
			    If(objConn.State = 0) Then ' Equals to If(objConn.State = 0)
			              Set objConn = Nothing
			              kyExcelDataExport = null
			              errHandler sColumnName
			              Exit Function
			    Else   
			              objRS.Open sSQLQuery, objConn, 3,3'Open the RecordSet by Executing the Query
			              errHandler sColumnName
			    End If	
			    	Set objConn = Nothing
					Set objRS = Nothing
			End If
End Function





Function kyExcelHeaderIndexFetch(sDataBase, sDataTable, sString)
on Error Resume Next
			If bKeywordPF = True Then
				strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBase & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
				sSQLQuery = "SELECT * FROM [" & sDataTable & "$]"
							Set objConn = CreateObject("ADODB.Connection")'Create Connection Object		
							Set objRS = CreateObject("ADODB.Recordset")'Create the RecordSet	
							objConn.ConnectionString = strConnection
						    objConn.ConnectionTimeOut = 10
						    objConn.CommandTimeOut = 10
			
			    objConn.Open' Open the Connection with DB mentioned in "strConnection"
			    
			    If(objConn.State = 0) Then ' Equals to If(objConn.State = 0)
			              Set objConn = Nothing
			              kySheetColumnNamesFetch = null
			              Exit Function
			    Else   
			              objRS.Open sSQLQuery, objConn, 3,3'Open the RecordSet by Executing the Query
			              iFieldsCount = objRS.Fields.Count
			              For iColCounter = 0 To iFieldsCount Step 1
			              	If Ucase(objRS.Fields(iColCounter).Name) = Ucase(sString) Then
			              		kyExcelHeaderIndexFetch = iColCounter
			              		Exit For
			              	Else
			              		kyExcelHeaderIndexFetch = null
			              	End If
			              Next
			
			    End If
			    	
			    	Set objConn = Nothing
					Set objRS = Nothing
					
		End If
End Function



Function kySheetColumnNamesFetch(sDataBase, sDataTable)
on Error Resume Next
			If bKeywordPF = True Then
					strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBase & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
					sSQLQuery = "SELECT * FROM [" & sDataTable & "$]"
								Set objConn = CreateObject("ADODB.Connection")'Create Connection Object		
								Set objRS = CreateObject("ADODB.Recordset")'Create the RecordSet	
								objConn.ConnectionString = strConnection
							    objConn.ConnectionTimeOut = 10
							    objConn.CommandTimeOut = 10
				
				    objConn.Open' Open the Connection with DB mentioned in "strConnection"
				    
				    If(objConn.State = 0) Then ' Equals to If(objConn.State = 0)
				              Set objConn = Nothing
				              kySheetColumnNamesFetch = null
				              Exit Function
				    Else   
				              objRS.Open sSQLQuery, objConn, 3,3'Open the RecordSet by Executing the Query
				              iFieldsCount = objRS.Fields.Count
				              ReDim arrFieldNames(iFieldsCount)
				              For iCount = 0 To iFieldsCount Step 1              	
						         arrFieldNames(iCount) = objRS.Fields(iCount).Name
				              Next
				              kySheetColumnNamesFetch = arrFieldNames
				    End If
				    	
				    	Set objConn = Nothing
						Set objRS = Nothing
						
			End If
End Function



Function kyGetSheetRecordset(sDataBase, sDataTable)
on Error Resume Next
			If bKeywordPF = True Then
					strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBase & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
					sSQLQuery = "SELECT * FROM [" & sDataTable & "$] where Execute = 'Yes'"
								Set objConn = CreateObject("ADODB.Connection")'Create Connection Object		
								Set objRS = CreateObject("ADODB.Recordset")'Create the RecordSet	
								objConn.ConnectionString = strConnection
							    objConn.ConnectionTimeOut = 10
							    objConn.CommandTimeOut = 10
				
				    objConn.Open' Open the Connection with DB mentioned in "strConnection"
				    
				    If(objConn.State = 0) Then ' Equals to If(objConn.State = 0)
				              Set objConn = Nothing
				              kyGetSheetRecordset = null
				              Exit Function
				    Else   
				              objRS.Open sSQLQuery, objConn, 3,3'Open the RecordSet by Executing the Query
				              iFieldsCount = objRS.Fields.Count
				              iRowsCount = objRS.RecordCount
				              ReDim arrDescKeyword(iRowsCount, iFieldsCount)
				                        		   
							              For iRowIterator = 0 To iRowsCount - 1 Step 1
							              		For iColIterator = 0 To iFieldsCount - 1 Step 1
							              			arrDescKeyword(iRowIterator,iColIterator) = objRS.Fields(iColIterator).Value
							              		Next
							              		
							              	objRS.MoveNext
							              Next   
								 
				              
				              
				              kyGetSheetRecordset = arrDescKeyword
				    End If
				    	
				    	Set objConn = Nothing
						Set objRS = Nothing
			End If
End Function



'**********************************************************************
'    Name: kyExcelDataInsert
'    Purpose:    This function inserts data into the Excel Sheet based on the provided parameters
'    Creator: Rajesh Ghosh
'
'        Param: sFullExcelPath| required 
'        AllowedRange: 
'        Description: Full path of the Excel workbook including the name.
'
'        Param: sSheetName| required 
'        AllowedRange: 
'        Description: Sheet Name where the data is to be entered

'
'        Param: sRowIdentifier
'        AllowedRange: 
'        Description: It is combination of two sets of Data in a row, On being matched, the date is written in same row.
'                    Syntax:    RowData1-ColumnNo|RowData2-ColumnNo
'
'        Param: sColumnIdentifier| required 
'        AllowedRange: 
'        Description: Column header name under which the data is to be entered

'        Param: sValue| required 
'        AllowedRange: 
'        Description: The value to be entered into the Excel sheet
'
'      Returns: NA

'**********************************************************************
Function kyExcelDataInsert(sFullExcelPath, sSheetName, sRowIdentifier, sColumnIdentifier, sValue)
            On Error Resume Next
            bKeywordPF = True        
            If bKeywordPF = True Then' If the Previous Keyword is Passed then Execute the Code - Else No Action
                            Dim objObject        
                            bKeywordPF = False                
                    Set oExcel = CreateObject("Excel.Application")    
                        oExcel.Workbooks.Open sFullExcelPath
                    Set oMysheet = oExcel.ActiveWorkbook.Worksheets(sSheetName)
                    
                    iRowsCount = oMysheet.usedRange.Rows.Count 
                    iColsCount = oMysheet.usedRange.Columns.Count
                    
                    sIdentifierArray = Split(Trim(sRowIdentifier),"|")
                    sRowFirstIdentifier = Trim(Split(Trim(sRowIdentifier),"|")(0))
                    iRefCol1 = Cint(Split(sRowFirstIdentifier,"-")(1))
                                If UBound(sIdentifierArray) = 1 Then
                                        sRowSecIdentifier = Trim(Split(Trim(sRowIdentifier),"|")(1))
                                        iRefCol2 = Cint(Split(sRowSecIdentifier,"-")(1))
                                End If
        
                    'Get the Column position of the Specified Column header
                    For iCol = 1 To iColsCount Step 1
                                If Trim(oMysheet.Cells(1,iCol).Value) = sColumnIdentifier Then
                                    bKeywordPF = True
                                    Exit For
                                Else
                                    bKeywordPF = False
                                End If
                    Next
                    
                    If bKeywordPF = False Then
                       sKeywordError = "Error in 'kyExcelDataInsert'.  "  & sColumnIdentifier & " Not Found in Excel Column Headers  " & Err.Description    
                    Else
                                                
                                For iRows = 2 To iRowsCount Step 1
                                        sFirstData = Trim(oMysheet.Cells(iRows,iRefCol1).Value)
                                    If sRowSecIdentifier <> "" Then
                                        sSecData = Trim(oMysheet.Cells(iRows,iRefCol2).Value)
                                    
                    '************************************ Check if the Row with both the data identifiers are present ***************
	                                            If sFirstData = Split(sRowFirstIdentifier,"-")(0) and  sSecData = Split(sRowSecIdentifier,"-")(0) Then
	                                                bKeywordPF = True
	                                                Exit For
	                                            Else
	                                                bKeywordPF = False
	                                                sKeywordError = "Error in 'kyExcelDataInsert'.  The given Combination of Row identifiers are not Found in Excel Rows " & Err.Description
	                                                                                    
	                                            End If
                                    ElseIf sFirstData = Split(sRowFirstIdentifier,"-")(0) Then
                                    		bKeywordPF = True
                                            Exit For
                                    Else
                                                bKeywordPF = False
                                                sKeywordError = "Error in 'kyExcelDataInsert'.  The given Combination of Row identifiers are not Found in Excel Rows " & Err.Description
									End If                                            
                                Next
                        
                                If isNumeric(sValue) Then
                                    oMysheet.Cells(iRows,iCol).Value = sValue
                                Else 
                                    oMysheet.Cells(iRows,iCol).Value = Cstr(sValue)
                                End If                    
                    End If    
                                        
                        oExcel.ActiveWorkbook.Save
                        oExcel.Application.Quit
'                        kyKillProcess "Excel.exe"
                        Set oMysheet = Nothing
                        Set oExcel =  Nothing
            End If 
End Function



Function kyExcelBulkDataInsert(EXEC, UNIVERSAL, sFullExcelPath, sSheetName, oExcelDict, sColumnIdentifier)
            On Error Resume Next
            bKeywordPF = True        
            If bKeywordPF = True Then' If the Previous Keyword is Passed then Execute the Code - Else No Action                


					asExcelVariables =  oExcelDict.Keys
				If lbound(asExcelVariables) >= 0 and asExcelVariables(lbound(asExcelVariables)) <> "" Then
					 	
					       
					Set oExcelSample = CreateObject("Scripting.Dictionary")   
					

                    Set oExcel = CreateObject("Excel.Application")    
                        oExcel.Workbooks.Open sFullExcelPath
                    Set oMysheet = oExcel.ActiveWorkbook.Worksheets(sSheetName)
                    
                    iRowsCount = oMysheet.usedRange.Rows.Count 
                    iColsCount = oMysheet.usedRange.Columns.Count
                    
                   
                    'Get the Column position of the Specified Column header
                    For iCol = 1 To iColsCount Step 1
                                If Trim(oMysheet.Cells(1,iCol).Value) = sColumnIdentifier Then
                                    bKeywordPF = True
                                    Exit For
                                Else
                                    bKeywordPF = False
                                End If
                    Next

						If bKeywordPF Then
											For iIterator = 2 To iRowsCount Step 1
												sTCName = oMysheet.Cells(iIterator,1)
												sVarName = oMysheet.Cells(iIterator,2)
																	If Ucase(sTCName) = Ucase(sTestCaseName) Then
																		oExcelSample.Add sVarName, iIterator
																	End If 			
											Next
			
			




											For iVarCount = Lbound(asExcelVariables) To Ubound(asExcelVariables) Step 1
															If oExcelSample.Exists(asExcelVariables(iVarCount)) Then
																iRowNum = oExcelSample.Item(asExcelVariables(iVarCount))
																sValue = oExcelDict.Item(asExcelVariables(iVarCount))
																		    If isNumeric(sValue) Then
											                                    oMysheet.Cells(iRowNum,iCol).Value = Cstr(sValue)
											                                Else 

											                                    oMysheet.Cells(iRowNum,iCol).Value = Cstr(sValue)
											                                End If  
															Else


																bKeywordPF = False
																sKeywordError = "The variable name: <b>"&asExcelVariables(iVarCount) & "</b> not present in the Excel"
																Exit For
															End If
											Next
									
						End If  

						oExcel.ActiveWorkbook.Save
                        oExcel.Application.Quit
                        Set oMysheet = Nothing
                        Set oExcel =  Nothing
                        oExcelSample.RemoveAll
			End If
                     
       End If
       
End Function




'**********************************************************************
'	Name: kyKillProcess
'	Purpose:This function kills all instances of a process.
'	Creator:Rajesh Ghosh
'
'		Param: sProcess| required
'		AllowedRange: 
'		Description: The process to kill.
'
'	Returns: True or False

'**********************************************************************
Public Sub kyKillProcess(sProcess)

	Dim oWMI
	Dim ret
	Dim sService
	Dim oWMIServices
	Dim oWMIService
	Dim oServices
	Dim oService
	Dim servicename
	 
	Set oWMI = GetObject("winmgmts:")
	Set oServices = oWMI.InstancesOf("win32_process")
	
	For Each oService In oServices
			 
		servicename = LCase(Trim(CStr(oService.Name) & ""))
			 
		If LCase(servicename) = LCase(sProcess) Then
			ret = oService.Terminate
		End If
	
	Next
	  
	Set oServices = Nothing
	Set oWMI = Nothing
   
End Sub
