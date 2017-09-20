
Function kySAPInstanceClose()
							
		SystemUtil.CloseProcessByName "saplogon.exe"

End Function




Function kyGetSAPObjectProperty(sTestObjectName, sPropertyType)
	On Error Resume Next
	If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		
				bKeywordPF = False
				Set objObject = kyGetSAPObject (sTestObjectName)' Get the Object with mentioned Description Name
				If (not isEmpty(objObject)) Then ' objObject <> Empty
					kyGetSAPObjectProperty = objObject.GetRoProperty(sPropertyType)
					errHandler sTestObjectName		
				Else
					kyGetSAPObjectProperty = ""		
				End If
		
	End If
End Function


Function kySAPTableFill(sTestObjectName, iRowID, sHeaderNames, sValues)

On Error Resume Next
	If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action

		bKeywordPF = False
		Set objObject = kyGetSAPObject (sTestObjectName)' Get the Object with mentioned Description Name
		If (not isEmpty(objObject)) Then ' objObject <> Empty
			asHeaderNames = Split(sHeaderNames, "|")
			asValues = Split(sValues, "|")
			If UBound(asHeaderNames) = Ubound(asValues) Then
				For iIterator = LBound(asHeaderNames) To  UBound(asHeaderNames) Step 1
					objObject.SetCellData iRowID, asHeaderNames(iIterator), asValues(iIterator)
						errHandler asHeaderNames(iIterator)
					If Environment("ErrorMsg") <> "" Then
						Exit Function
					End If
				Next
				
			Else
				Environment("ErrorMsg") = "Values count mismatch with the fields count.<br>"&_
										  "Please check the Table fields and its respective values count are same."
			End If
		
		End If
	
	End If
End Function




Function kyGetSAPObject (sTestObjectName)

		On Error Resume Next	
			Dim sBrowserName, sPageTitle, objAppObject, sObjectClass, objPropertyCollection, sPropertyValuePairs
			objAppObject = NULL
		
		if(InStr(sTestObjectName,";;") > 0) Then
					sPropertyValuePairs = Split(sTestObjectName,";;")(1)
					sObjectClass = Split(sTestObjectName,";;")(0)
					Set objPropertyCollection = kyGetPropertiesCollection((sPropertyValuePairs))
					wait 10

					Execute "Set  objAppObject = " & sObjectClass & "(objPropertyCollection)"
		Else
					Execute "Set  objAppObject = " & sTestObjectName
		End If
			
		If objAppObject.Exist(synTime) <>  Empty OR  objAppObject.Exist(synTime) = True Then
					Set kyGetSAPObject = objAppObject						
		Else
					Set kyGetSAPObject = empty 
					fLibTempPopup "Message from Funciton kyGetSAPObject", sTestObjectName & " Does NOT Exist"
					arrObject = Split(sTestObjectName, ".")
					Environment("ErrorMsg") = "<b>"&arrObject(Ubound(arrObject)) &"</b><br>"&_
												"Object not found on page.<br>"&_
						 						"Please check the object properties.<br>"					
		End If
			
End Function



Function kyGetPropertiesCollection (sObjectsProperties)

Dim objPropColl, arrProperties, arrValues(), iPropertiesCount, arrObjectsPropertiesAndValues
arrObjectsPropertiesAndValues = Split(sObjectsProperties,",")
Set objPropColl = Description.Create()

ReDim arrProperties(UBound(arrObjectsPropertiesAndValues))
ReDim arrValues(UBound(arrObjectsPropertiesAndValues))

		For iPropertiesCount = 0 to UBound(arrObjectsPropertiesAndValues)
				arrProperties(iPropertiesCount)  = Split(arrObjectsPropertiesAndValues(iPropertiesCount),":=")(0)
				arrValues(iPropertiesCount) 		= Split(arrObjectsPropertiesAndValues(iPropertiesCount),":=")(1)
		Next

	For iPropertiesCount = 0 to UBound(arrProperties)
		objPropColl(arrProperties(iPropertiesCount)).Value = arrValues(iPropertiesCount)
	Next

Set kyGetPropertiesCollection = objPropColl

End Function



Function kyPerformSAPOperation (sTestObjectName, sAction, sText)
	On Error Resume Next
	If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		Dim objObject, sTextArray, sTextArrayLength, executeText
		bKeywordPF = False
		Set objObject = kyGetSAPObject (sTestObjectName)' Get the Object with mentioned Description Name
		If (not isEmpty(objObject)) Then ' objObject <> Empty
			classDescription = objObject.GetTOProperty("Class Name")
			If sText <> "" AND InStr(sText, ";") <> 0 Then
'				sTextArray = Split(sText, ";")
'				sTextArrayLength = UBound(sTextArray)
'				executeText = "objObject."& sAction  & " "				
'				For i = 0 To sTextArrayLength
'					If i = 0 AND i = sTextArrayLength Then
'						executeText = executeText & "" & sTextArray(i) &""""
'					ElseIf i = 0 AND i <> sTextArrayLength Then
'						executeText = executeText & "" & sTextArray(i) &""
'					ElseIf i > 0 AND i <= sTextArrayLength Then
'						executeText = executeText & ", """ & sTextArray(i) &""""
'					Else
'						executeText = "objObject." & sAction & " """ & sTextArray(i) & """"
'					End If	
'				Next
				executeText = "objObject." & sAction & " """ & sText & """"	
			ElseIf sText <> "" AND not InStr(sText, ";") <> 0 Then
				If IsNumeric(sText) and classDescription <> "OracleTextField" Then

					executeText = "objObject." & sAction & " "&sText
				ElseIf classDescription = "OracleTextField" and objObject.GetRoProperty("editable") = False Then
					sAction = "nothing"	
				Else
					executeText = "objObject." & sAction & " """ & sText & """"				
				End If
			Else
				executeText = "objObject." & sAction
			End If
			
			'Finally Execute the action on the object
					If sAction <> "nothing" Then
							Execute executeText
							errHandler sTestObjectName
					End If
		Else
				

		End If
	End If	 
		
End Function


Function kySAPWaitTillObjectExist(sObjectName, sTime)
	On Error Resume Next		
		If bKeywordPF = True Then	
					Dim objAppObject
					
					Execute "Set  objAppObject = " & sObjectName
					dtTime = DateAdd("s",sTime,now)
					While objAppObject.Exist(1) = False and dtTime > now
						
					Wend
					If objAppObject.Exist(1) Then
							bKeywordPF  =True
					Else
							bKeywordPF = False
						sError = "Error in 'kySAPWaitTillObjectExist'. Object  " & CStr(sObjectName) & "  NOT Found on Page"	
					End If
					
'				End If 
		End If 
		kySAPWaitTillObjectExist = bKeywordPF					
End Function


Function listItemSelect (sObjectName, listItemToSelect)
	On Error Resume Next
	If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		Dim objObject, sTextArray, sTextArrayLength, executeText
		bKeywordPF = False
		Set listObject = kyGetSAPObject (sObjectName)' Get the Object with mentioned Description Name
		classDescription = objObject.GetTOProperty("Class Name")
		If (not isEmpty(listObject)) Then ' objObject <> Empty
						listItemSelect = false
						listItemsCount = listObject.GetItemsCount
						listNumericItemExtraText = ""
					
					'	numeric value
						If isNumeric(listItemToSelect) = true Then
					
						'	is item pozition within items count limit?
							If listItemToSelect => listItemsCount Then
								enterReport"RegularTextReport", "failed", "List item <strong>" & listItemToSelect &" </strong> was not possible to select since there is less rows or no row (rows count: <strong>" & listItemsCount & "</strong>)." 			
							Else
								listObject.Select(listItemToSelect)
								itemVal = Trim(listObject.GetItem (listItemToSelect))
								If itemVal <> "" Then
									listNumericItemExtraText = "("&itemVal&")"
								End If
								listItemSelect = true
							End If
					                   
					'	text value
						Else
						
							allListItems = Split(listObject.GetROProperty("all items"), chr(10))
							For i = 0 To UBound(allListItems)
								if allListItems(i) = listItemToSelect then
									listObject.RefreshObject
									listObject.Select(listItemToSelect)
									listItemSelect = true
									Exit For
								End if
							Next
							
						End If
			End IF

	End If
End Function

Function kySAPCheckBoxToggle (sTestObjectName, sStatus)
	On Error Resume Next							
				If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
							Set objObject = kyGetSAPObject (sTestObjectName)' Get the Object with mentioned Description Properties
							If (not isEmpty(objObject)) Then									
									value = objObject.GetROProperty("selected")		
			
											Select Case UCase(Trim(sStatus))
												Case "CHECK"
													If value = False Then
														kyPerformSAPOperation sTestObjectName, "Set", "ON"
													End If
			
												Case "UNCHECK"
													If value = True Then
														kyPerformSAPOperation sTestObjectName, "Set", "OFF"
													End If
											End Select
									If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
									Else
												sKeywordError = "Error in 'kySAPCheckBoxToggle'  " & Err.Description	
												bKeywordPF = False
									End If
								
							 Else 
							 	sKeywordError = "Error in 'kySAPCheckBoxToggle'. Object  " & CStr(sTestObjectName) & "  NOT Found on Page"	
							 End IF	
				 End IF	
End Function


Function kyFillTableRowByRowAndColID(sTestObjectName, iRowId, iColId, sCellData)
		On Error Resume Next							
				If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
							bKeywordPF = False
							Set objObject = kyGetSAPObject (sTestObjectName)' Get the Object with mentioned Description Properties
							If (not isEmpty(objObject)) Then
								objObject.SetCellData iRowId, iColId, sCellData
								If Err.Number = 0 Then
									bKeywordPF = True
								End If
							End If				
					
				End If
End Function
	
	
	
Function kySAPIfObjectExistsCheck(sObjectsProperties, sTime)

	On Error Resume Next		
	If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
				Dim objObject
				
				Execute "Set objObject = "&sObjectsProperties

				
				dtTime = DateAdd("s",sTime,now)
				While objObject.Exist(0)=False and dtTime > now 			

				Wend	
				
				If objObject.Exist(0) Then
					kySAPIfObjectExistsCheck = True
				Else 	
					kySAPIfObjectExistsCheck = False
				End If
	End IF

	 
End Function


Sub kyStatusBarSync ()
			On Error Resume Next
			If bKeywordPF = True Then
					kyPerformSAPOperation OBJ_Comm_SAPGuiStatusBar_Window_StatusBar, "Sync", ""
			End If
End Sub


Sub kySAPSync (sTestObject, waitTime)
			On Error Resume Next
			If bKeywordPF = True Then
					Execute "Set objObject = "&sObjectsProperties
					objObject.Sync
					If  waitTime =  "" Then
						Wait 3
					Else
						Wait  waitTime
					End If
			End If
End Sub


Function kySAPGetObjectProperty(sTestObjectName, sObjProperty)
			On Error Resume Next		
			If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
						Dim objObject
						bKeywordPF = False
						Set objObject = kyGetSAPObject (sTestObjectName)' Get the Object with mentioned Description Properties
						If (not isEmpty(objObject)) Then
							kySAPGetObjectProperty = objObject.GetROProperty(sObjProperty)
								If Err.Number = 0 AND Err.Description = ""  Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kySAPGetObjectProperty'  " & Err.Description	
								End If
						Else
										sKeywordError = "Error in 'kySAPGetObjectProperty'. Object  " & CStr(sObjectsProperties) & "  NOT Found on Page"	
						End If
			End If 

End Function

Function kySAPCompareObjectProperty(sTestObjectName,sObjProperty,sExpValue)
				On Error Resume Next		
				If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
							bKeywordPF = False
							Set objObject = kyGetSAPObject (sTestObjectName)' Get the Object with mentioned Description Properties
							If (not isEmpty(objObject)) Then
								value = CStr(objObject.GetROProperty(sObjProperty))
								errHandler sTestObjectName
								If (CStr(value) = sExpValue) Then
									bKeywordPf = true
								Else
										arrObject = Split(sObj, ".")
										Environment("ErrorMsg") = "Property value doesn't match for the object<b>"&arrObject(Ubound(arrObject))&_
																  "</b><br>Please the re-check Actual and the Expected value. "										
								End If
							End if 
				End If 
	
End Function



