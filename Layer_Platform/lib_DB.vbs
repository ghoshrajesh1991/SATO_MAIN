On Error Resume Next
'''''''''''''''''''''''Function mInitialSetUp(UNIVERSAL)
'''''''''''''''''''''''				
'''''''''''''''''''''''				arrLoginData = "http://10.83.17.43:7001/lis/welcome.do,patilked,patilked,Private customer"
'''''''''''''''''''''''				arrTaskData = "Private customer,,,,M,Kedarnath,Patil,01.01.1980,English,Finland,,,,,12,Regular peyment,Invoice,2000,1000,12,01.01.2020,28"
'''''''''''''''''''''''				arrTaskData1 = "Private customer,,,,M,Kedarnath,Patil,01.01.1980,English,Finland,Finland,Eura,Broker,Family member"				
'''''''''''''''''''''''				arrTaskData2 = "Private customer,,,,M,Kedarnath,Patil,01.01.1980,English,Finland,Finland,Eura,Broker,Family member"
'''''''''''''''''''''''				
'''''''''''''''''''''''				arrLoginData = Split(arrLoginData,",")
'''''''''''''''''''''''				arrTaskData = Split(arrTaskData,",")
'''''''''''''''''''''''
'''''''''''''''''''''''				UNIVERSAL.Item("APPURL") = arrLoginData(0)'"http://10.83.17.43:7001/lis/welcome.do"
'''''''''''''''''''''''				UNIVERSAL.Item("USERNAME") = arrLoginData(1)'"patilked"
'''''''''''''''''''''''				UNIVERSAL.Item("PASSWORD") = arrLoginData(2)'"patilked"
'''''''''''''''''''''''
'''''''''''''''''''''''				UNIVERSAL.item("CUSTOMERTYPE") = arrTaskData(0)'"Private customer"
'''''''''''''''''''''''				UNIVERSAL.item("PERSONALID") = arrTaskData(1)'""
'''''''''''''''''''''''				UNIVERSAL.item("BUSINESSID") = arrTaskData(2)'""
'''''''''''''''''''''''				UNIVERSAL.item("CUSTCHANNELCODE") = arrTaskData(3)'""
'''''''''''''''''''''''
'''''''''''''''''''''''				UNIVERSAL.item("CUSTOMERGENDER") = arrTaskData(4)'"M"
'''''''''''''''''''''''				UNIVERSAL.item("FIRSTNAME") = arrTaskData(5)'"Kedarnath"
'''''''''''''''''''''''				UNIVERSAL.item("LASTNAME") = arrTaskData(6)'"Patil"
'''''''''''''''''''''''				UNIVERSAL.item("BIRTHDATE") = arrTaskData(7)'"01.01.1980"
'''''''''''''''''''''''				UNIVERSAL.item("CUSTLANGUAGE") = arrTaskData(8)'"English"
'''''''''''''''''''''''				UNIVERSAL.item("CUSTNATIONALITY") = arrTaskData(9)'"Finland"
'''''''''''''''''''''''				UNIVERSAL.item("CUSTCOUNTRY") = arrTaskData(10)'"English"
'''''''''''''''''''''''				UNIVERSAL.item("CUSTTAXMUNICIPALITY") = arrTaskData(11)'"English"
'''''''''''''''''''''''				UNIVERSAL.item("SOURCE") = arrTaskData(12)'"English"
'''''''''''''''''''''''				UNIVERSAL.item("NONCUSTSTATUS") = arrTaskData(13)'"English"
'''''''''''''''''''''''
'''''''''''''''''''''''				UNIVERSAL.item("PRODUCT") = arrTaskData(14)
'''''''''''''''''''''''
'''''''''''''''''''''''				UNIVERSAL.item("PAYMENTPLANTYPE") = arrTaskData(15)
'''''''''''''''''''''''				UNIVERSAL.item("PAYMENTMETHOD") = arrTaskData(16)
'''''''''''''''''''''''				UNIVERSAL.item("FIRSTPREMIUMAMOUNT") = arrTaskData(17)
'''''''''''''''''''''''				UNIVERSAL.item("PREMIUMAMOUNT") = arrTaskData(18)
'''''''''''''''''''''''				UNIVERSAL.item("PAYMENTSPERYEAR") = arrTaskData(19)
'''''''''''''''''''''''				UNIVERSAL.item("PAYMENTPERIODENDDATE") = arrTaskData(20)
'''''''''''''''''''''''				UNIVERSAL.item("PAYMENTDAY") = arrTaskData(21)
'''''''''''''''''''''''				
'''''''''''''''''''''''				'msgbox UNIVERSAL.item("FIRSTNAME") & " - - -" & UNIVERSAL.item("LASTNAME") & "- - - "&UNIVERSAL.item("BIRTHDATE") 
'''''''''''''''''''''''				'UNIVERSAL.RemoveAll
'''''''''''''''''''''''				'msgbox "SHREE"
'''''''''''''''''''''''
'''''''''''''''''''''''End Function

Function mFillUniversal(UNIVERSAL,TCID)

		 Dim arrTestData, strConnection, strSQL
		
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Pravin_QTP\TestData\TestData.xls;Extended Properties=""Excel 8.0;HDR=Yes;"";"
		strSQL = "SELECT Desc,Val FROM [Table1$] where TCID = "+TCID+""  
		'msgBox strSQL
		arrTestData = mGetMultipleData(strConnection, strSQL) ' Fetch the Test Data for Specific Test case
		For iDataCount = 0 to UBound(arrTestData)
				UNIVERSAL.Add arrTestData(iDataCount, 0), arrTestData(iDataCount, 1)
		Next

End Function

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


Sub  mUpdateLogFile(TCID)

		On Error Resume Next
		Dim  objConn,  sSQLQuery, sErrorText1

		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Pravin_QTP\TestResultLogs\T14.S529.xls;Extended Properties=""Excel 8.0;HDR=Yes;"";"
	    sSQLQuery =  "UPDATE [Table1$] SET [Status] = 'Pass', [Status Date] = '2010', [Status Remark]='Test Case is Passed', [Status Snap Shot Link]='No Link' where [TCID] = " & TCID & ""
   ' strConnection = "Driver=Microsoft Excel Driver (*.xls);DBQ=C:\myData.xls;Extended Properties=""Excel 8.0;HDR=Yes;"";"
    
    Set objConn = CreateObject("ADODB.Connection")'Create Connection Object

    objConn.Open strConnection' Open the Connection with DB mentioned in "strConnection"
    If(objConn.State = 0) Then ' Equals to If(objConn.State = 0)
				  sErrorText1 = mGetConnectionErrors (objConn, "'mUpdateLogFile' at DB Connection")
				  objConn.Close
				  Set objConn = Nothing
				   sLogDataToFile "C:\Pravin_QTP\TestResultLogs\Logs\DBLogs.txt", mGetTimeStampLong & "--- " & sErrorText1
				  Exit Sub
    Else   
				 objConn.Execute sSQLQuery
				 sLogDataToFile "C:\Pravin_QTP\TestResultLogs\Logs\DBLogs.txt", mGetTimeStampLong & "--- Connected to DB Successfully. Connection String:- " &  strConnection 
				  If objConn.Errors.Count > 0 Then' Equals to If(objConn.State = 0)
						  sErrorText1 = mGetConnectionErrors (objConn, "'mUpdateLogFile' at RecordSet Query")
						  MsgBox sErrorText1
						  objConn.Close
						  Set objConn = Nothing
						  sLogDataToFile "C:\Pravin_QTP\TestResultLogs\Logs\DBLogs.txt", mGetTimeStampLong & "--- " & sErrorText1
						  Exit Sub
				  Else
							sLogDataToFile "C:\Pravin_QTP\TestResultLogs\Logs\DBLogs.txt", mGetTimeStampLong & "--- Executed the Query Successfully:- " & sSQLQuery			  
				  End If ' Enf of If Statement for Verifying the RecordSet Record Count
						  objConn.Close
						  Set objConn = Nothing                      
  
    End If ' End of If Statement for Verifying Connection Error
End Sub




Sub te
   On Error Resume Next
   Dim  objConn, objRS,  sSQLQuery,TCID
TCID = "112"
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Pravin_QTP\TestResultLogs\T14.S529.xls;Extended Properties=""Excel 8.0;HDR=Yes;"";"
	    sSQLQuery =  "UPDATE [Table1$] SET [Status] = 'Pass', [Status Date] = '2010', [Status Remark]='Test Case is Passed', [Status Snap Shot Link]='No Link' where [TCID] = " & TCID & ""
   ' strConnection = "Driver=Microsoft Excel Driver (*.xls);DBQ=C:\myData.xls;Extended Properties=""Excel 8.0;HDR=Yes;"";"
    
	'strSQL = "SELECT Desc,DescVal FROM [Table1$]"  
'    MsgBox sSQLQuery
    Set objConn = CreateObject("ADODB.Connection")'Create Connection Object
    objConn.ConnectionString = strConnection
    objConn.ConnectionTimeOut = 10
    objConn.CommandTimeOut = 10
objConn.Open
' Set objRS = CreateObject("ADODB.Recordset")
' msgbox (err.description())
' objRS.Open sSQLQuery, objConn
objConn.Execute sSQLQuery
 if objConn.Errors.Count > 0 Then

 msgbox ("Error Number:  " & objConn.Errors.Item(0).Number  & "     Error Message: " & objConn.Errors.Item(0).Description)

 End if
End Sub