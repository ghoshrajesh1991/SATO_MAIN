
'=======================================================================

' This File/Library contains the Ovject level functions for the "Database Interactions" - xls files

'List of functions in the File
'		mGetConnectionErrors - 			 	by - Swapnali
'		mGetMultipleData - 			 			 by - Swapnali
'		subUpdateResultFile - 					by - Swapnali

'=======================================================================
On Error Resume Next

Function mGetConnectionErrors( objConn, sErrorAt)
		
      Dim arrConnectionErrors(), iConnectionErrorCount, iCount, sErrorText
      sErrorText = ""
      iConnectionErrorCount = objConn.Errors.Count
      ReDim arrConnectionErrors(iConnectionErrorCount)
      For iCount = 0 To  iConnectionErrorCount - 1
            arrConnectionErrors(iCount) = objConn.Errors.Item(iCount).Description
            sErrorText = sErrorText & "  " & CStr(iCount) & "  "  & arrConnectionErrors(iCount)  
      Next
    
      mGetConnectionErrors = "Error in " & sErrorAt & "  " & sErrorText
   
End Function

Function mGetMultipleData(strConnection, sSQLQuery)
    'MsgBox "Shree"
		Dim  objConn, objRS, arrRowsData(), objField, iExtArrayLength, iIntArrayLength, sErrorText1
	   
   ' strConnection = "Driver=Microsoft Excel Driver (*.xls);DBQ=C:\myData.xls;Extended Properties=""Excel 8.0;HDR=Yes;"";"
    
	'strSQL = "SELECT Desc,DescVal FROM [Table1$]"  
    
    Set objConn = CreateObject("ADODB.Connection")'Create Connection Object
    objConn.ConnectionString = strConnection
    objConn.ConnectionTimeOut = 10
    objConn.CommandTimeOut = 10

    objConn.Open' Open the Connection with DB mentioned in "strConnection"
    If(objConn.State = 0) Then ' Equals to If(objConn.State = 0)
			  sErrorText1 = ""
			  sErrorText1 = mGetConnectionErrors (objConn, "'mGetMultipleData' at DB Connection")
              ReDim arrRowsData(0,0)
              arrRowsData(0,0) = "FALSE"
              objConn.Close
              Set objConn = Nothing
              Exit Function
    Else   
              Set objRS = CreateObject("ADODB.Recordset")'Create the RecordSet
              objRS.Open sSQLQuery, objConn, 3,3'Open the RecordSet by Executing the Query
              
              If(objRS.State = 0) Then ' Equals to If(objConn.State = 0)
					  sErrorText1 = ""
                      sErrorText1 = mGetConnectionErrors (objConn, "'mGetMultipleData' at RecordSet Query")
                      arrRowsData(0,0) = "FALSE"
                       objRS.Close
                      Set objRS = Nothing
                      objConn.Close
                      Set objConn = Nothing
                      'MsgBox " Record State :-  " & objRS.State & "----- RecordSet Errors Count:-  " & objConn.Errors.Count , ,"SHREE"
                      Exit Function
              Else               
                      'MsgBox " Record State :-  " & objRS.State & "----- RecordSet Records Count:-  " & objRS.RecordCount & "----- RecordSet Fields Count:-  " & objRS.Fields.Count
                      If(objRS.RecordCount < 0) Then ' Check if the RecordCount is greater than Zero
                              ReDim arrRowsData(0,0)
                              arrRowsData(0,0)  = "NULL"
                      Else
                              iIntArrayLength = 0
							  iIntArrayLength = objRS.Fields.Count
                              ReDim arrRowsData(objRS.RecordCount - 1, iIntArrayLength - 1) 
                              iExtArrayLength = 0
							  objRS.MoveFirst
                              WHILE NOT objRS.EOF
                                      iIntArrayLength = 0  
                                      For Each objField In objRS.Fields
                                              arrRowsData(iExtArrayLength, iIntArrayLength) =  objField.Value
                                              iIntArrayLength = iIntArrayLength + 1
                                      Next                              
                                      'MsgBox objRS.Fields("Name") & "-- " & objRS.Fields("Surname")
                                      objRS.MoveNext
                                      iExtArrayLength = iExtArrayLength + 1
                              WEND                                      
                      End If ' Enf of If Statement for Verifying the RecordSet Record Count
					  mGetMultipleData = arrRowsData
                      objRS.Close
                      Set objRS = Nothing
                      objConn.Close
                      Set objConn = Nothing                      
              End If ' End of If Statement for Verifying RecordSet Error    
    End If ' End of If Statement for Verifying Connection Error
End Function


Sub  subUpdateResultFile(UNIVERSAL,EXEC, sDataBase, sDataTable, TCID, sStatus, sStatusRemark,sTestFlow)

		On Error Resume Next
		Dim  objConn,  sSQLQuery, sErrorText1, sSnapShotLink, objFSO

		'EXEC.Item("TCFOLDER") = EXEC.Item("TASKRESULTFOLDER") & "\" & TCID

		Set objFSO= CreateObject("Scripting.FileSystemObject")
		Set TaskResultFolder= objFSO.CreateFolder(EXEC.Item("TCFOLDER"))

		'If sStatus = "Fail" Then

			Browser(Environment("sBrowser")).SetTOProperty "CreationTime", iBrowserCreationTimeGlobal
'			sSnapShotLink = EXEC.Item("TASKRESULTFOLDER")  & EXEC.Item("TASKRESULTFILE")
			sSnapShotLink = EXEC.Item("TCFOLDER")  & "\" & TCID
'''''			'Browser("name:=.*").CaptureBitmap(sSnapShotLink)
			'CaptureSnapshot Browser(Environment("sBrowser")).Page(Environment("sPage")),  sSnapShotLink
			CaptureSnapshot_new  Browser(Environment("sBrowser")).Page(Environment("sPage")),  sSnapShotLink

		'End If
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBase & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
		sStatusRemark = Replace (sStatusRemark,"'",":")
	 sSQLQuery =  "UPDATE [" & sDataTable & "$] SET [Status] = '" & sStatus & "', [TimeStamp] = '" & mGetTimeStampLong() & "', [Remark]='" & Left(sStatusRemark,256) & "',[TestFlow]='" &sTestFlow& "' , [SnapShots]='" & sSnapShotLink & "' where [TCID] = '" & TCID & "'"
		'sSQLQuery =  "UPDATE [" & sDataTable & "$] SET [Status] = '" & sStatus & "', [TimeStamp] = '" & mGetTimeStampLong() & "', [Remark]=""" & Left(sStatusRemark,256) & """, [TestFlow]='" &UNIVERSAL.Item("TESTFLOW")& "' ,[SnapShots]='file://" & sSnapShotLink & ".html' where [TCID] = '" & TCID & "'"
   ' strConnection = "Driver=Microsoft Excel Driver (*.xls);DBQ=C:\myData.xls;Extended Properties=""Excel 8.0;HDR=Yes;"";"
'    MsgBox sSQLQuery
		fLibTempPopup "Query at Sub 'subUpdateResultFile'  ", sSQLQuery
    Set objConn = CreateObject("ADODB.Connection")'Create Connection Object

    objConn.Open strConnection' Open the Connection with DB mentioned in "strConnection"
    If(objConn.State = 0) Then ' Equals to If(objConn.State = 0)
				  sErrorText1 = mGetConnectionErrors (objConn, "'subUpdateResultFile' at DB Connection")
				  objConn.Close
				  Set objConn = Nothing
				   subLogDataToFile cTESTRESULTLOGFOLDERPATH & "Logs\DBLogs.txt", mGetTimeStampLong & "******* " & sErrorText1
				  Exit Sub
    Else   
				 objConn.Execute sSQLQuery
				 subLogDataToFile cTESTRESULTLOGFOLDERPATH & "Logs\DBLogs.txt", mGetTimeStampLong & "------- Connected to DB Successfully. Connection String:- " &  strConnection 
				  If objConn.Errors.Count > 0 Then' Equals to If(objConn.State = 0)
						  sErrorText1 = mGetConnectionErrors (objConn, "'mUpdateLogFile' at RecordSet Query")
						  'MsgBox sErrorText1
						  objConn.Close
						  Set objConn = Nothing
						  subLogDataToFile cTESTRESULTLOGFOLDERPATH & "Logs\DBLogs.txt", mGetTimeStampLong & "******* " & sErrorText1
						  Exit Sub
				  Else
							subLogDataToFile cTESTRESULTLOGFOLDERPATH & "Logs\DBLogs.txt", mGetTimeStampLong & "------- Executed the Query Successfully:- " & sSQLQuery			  
				  End If ' Enf of If Statement for Verifying the RecordSet Record Count
						  objConn.Close
						  Set objConn = Nothing                      
  
    End If ' End of If Statement for Verifying Connection Error
End Sub
