'=======================================================================

' This File/Library contains the Ovject level functions for the "Web Application"

'List of functions in the File
'		kyIfObjectExists - 										
'		kyIfObjectEnabled - 								
'		kyIfObjectDisabled - 								
'		kyRowWithCellText - 							
'		kyGetObjectProperty - 							
'		kyGetPropertiesCollection - 				
'		kyGetObject - 									  		  
'		kyPerformWebOperation - 		  		  
'		kyActionOnSpecificTableCell - 	 		 
'		kyIfTableContainsSpecifiedRows - 	 
'		kyIfObjectExistsShortWait - 				
'		kyIfObjectNOTOnPageShortWait - 	   
'		kyGetTableCellData - 							    
'		kyCompareTableCellData - 					 
'		kyWaitTillObjectTextChange - 			  
'		kyClickObjTillItExsists - 						    
'		kyCompareObjectProperty - 				   
'		kyNavigation - 							   				   

'=========================================================================================================================================================

' This Function returns the Description Object Based on the Inputs
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- Test Object with Hirarchy from Broser > Page > Object

Function kyGetObject1 (sTestObjectName)

			On Error Resume Next
			
			Dim objAppObject, sObjectClass, objPropertyCollection, sPropertyValuePairs
			Dim className
			objAppObject = NULL
			
		
			if(JavaMenuName=(Split(sTestObjectName,"_")(0) &"_"))then  ' this if is for menu items only
				Level=0			
			Else				
				Level=Split(sTestObjectName,"_")(0)			
			End if



			Select Case Level
					Case 0
						finalExecuteString=getMenuExcuteString(sTestObjectName)		
						Execute "Set  objAppObject =" & finalExecuteString	
						wait w2
						
					Case 1
						
						sObjectClass = Split(sTestObjectName,"_")(1) 
						Execute "Set  objAppObject = JavaWindow(win01)." & sObjectClass & "(""" & sTestObjectName & """)"
							
					Case 2
						
						sObjectClass = Split(sTestObjectName,"_")(2)
						WinnName=Split(sTestObjectName,"_")(1)
												
						'added for dialog window and objects
						if (ubound (Split(sTestObjectName,"_"))=4)then
							objectLocation=Split(sTestObjectName,"_")(4)  'OnDialog
						End if 	
						
						If objectLocation="OnDialog" Then
							Execute "Set  objAppObject = JavaWindow(win02).JavaDialog(WinnName)." & sObjectClass & "(""" & sTestObjectName & """)"
						Else
							Execute "Set  objAppObject = JavaWindow(win02).JavaInternalFrame(WinnName)." & sObjectClass & "(""" & sTestObjectName & """)"
						End If
						
					Case 3
						
						sObjectClass = Split(sTestObjectName,"_")(1)
						Execute "Set  objAppObject = JavaWindow(win03)." & sObjectClass & "(""" & sTestObjectName & """)"					

					Case 4
						
						sObjectClass = Split(sTestObjectName,"_")(2)
						WinnName=Split(sTestObjectName,"_")(1)
						'TeWindow("TeWindow").TeTextScreen("TeTextScreen").Type micShiftDwn +  micF8  + micShiftUp
						Execute "Set  objAppObject = TeWindow(win04).TeScreen(WinnName)." & sObjectClass & "(""" & sTestObjectName & """)"
					Case 5
					
						sObjectClass = Split(sTestObjectName,"_")(1)
						Execute "Set  objAppObject = Dialog(win05)." & sObjectClass & "(""" & sTestObjectName & """)"	
						'Dialog("Connect to Host - TN3270").WinEdit("Edit").SetSelection 0,6
						'Dialog("Connect to Host - TN3270").WinEdit("Edit_2").SetSelection 0,6
						'Dialog("TN3270 Plus").WinButton("OK").Click
						'Dialog("Connect to Host - TN3270").WinRadioButton("Telnet").Set
						'Dialog("Connect to Host - TN3270").WinButton("Connect").Click

					Case Else 
		
			End Select

			If objAppObject.Exist(5) = Empty OR  objAppObject.Exist(5) = False Then
						Set kyGetObject = kyGetDummyObject(sObjectClass)
						fLibTempPopup "Message from Funciton kyGetObject", sTestObjectName & " Does NOT Exist on Page"
						
			Else
						Set kyGetObject = objAppObject				
			End If



End Function


'=========================================================================================================================================================

' This Function Performs the desired Operation on the Web Object
'Input:-  Object Type, Objects Logical Name in the Shared Repository, Action to Perform, Value to be entered-selected if any
'Output:- Complete Test Object with Hirarchy from Broser

		Function kyPerformWinOperation ( sTestObjectName, sAction, sText)
					On Error Resume Next
					If bKeywordPF = True AND ( sAction = "Click" OR sAction = "SelectMenuItem" OR sAction = "Refresh" OR sText <> Empty) Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action  '
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Name
								If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty
								
								
										Select Case sAction
										
											Case "Select"
													objObject.Select(sText)													
											
											Case "Click"
													Execute "objObject." & sAction		
													
											Case "Set"		
													Execute "objObject." & sAction & " """ & sText & """"
											Case "PressKey"
													wait w1
													Execute "objObject.PressKey(sText)"
	
											Case "SelectMenuItem"

													Execute "objObject.RefreshObject"
													wait w2
													Execute "objObject.Select"
													
											Case "Refresh"
													
													Execute "objObject.RefreshObject"
													wait w1		
													
											Case "Type"

	'												Set oDesc = Description.Create
	'												oDesc("to_class").value = "JavaEdit"
	'												Set obj =objObject.ChildObjects(oDesc)
	'												'Execute "obj.Set" & " """ & sText & """"
	'												Msgbox "Total Obj: " & obj.Count
	'												Execute "obj.Type(sText)"
	
													
													Execute "objObject.RefreshObject"
													Execute "objObject.DblClick 10,10,""LEFT"""	
													wait w1													
													Set WshShell = CreateObject("WScript.Shell")
													WshShell.SendKeys sText	
													
'													Execute "objObject.ChildObjects.Set" & " """ & sText & """"
'													wait w2
'													'Execute "objObject.PressKey(sText)"
'													'Execute "objObject.Click"
													
													'Execute "objObject.Type(sText)"		

													wait w1
													
										Case "TypeInTerminal"
													Execute "objObject.RefreshObject"
													Execute "objObject.Set(sText)"
																									
'													Set WshShell = CreateObject("WScript.Shell")
'													WshShell.SendKeys sText	
													wait w1	
										
										Case Else 
		
										End Select

										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyPerformWinOperation'  " & Err.Description	
										End If
								Else
												sKeywordError = "Error in 'kyPerformWinOperation'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
								End If
					End If	 
		
		End Function


'=========================================================================================================================================================
' This Function Verifies if the Object is "Present on Page and Enabled"
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- None. But it sets the value of "bKeywordPF" to "False" if the Object is absent or Disabled.

		Function kyIfObjectExists( sTestObjectName)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist) Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												kyIfObjectExists=true
										Else
												sKeywordError = "Error in 'kyIfObjectExists'  " & Err.Description
												kyIfObjectExists=false												
										End If
								Else
												sKeywordError = "Error in 'kyIfObjectExists'. Object  " & CStr(sTestObjectName) & "  NOT Found."	
								End If
					End If 
		
		End Function
'=========================================================================================================================================================

		Function kyActivateAppWindow()
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								wait w1
									JavaWindow(win02).JavaObject("aQ").dblClick 30,7,"LEFT"
								wait 1
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												kyActivateAppWindow=true
										Else
												sKeywordError = "Error in 'kyActivateAppWindow'  " & Err.Description
												kyActivateAppWindow=false												
										End If
								Else

					End If 
		
		End Function
		
'=========================================================================================================================================================

		Function kyActivateTerminal()
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								wait w1
								TeWindow("TeWindow").WinMenu("Menu").Select "Edit;Copy	Ctrl+C"
								wait 1
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												kyActivateTerminal=true
										Else
												sKeywordError = "Error in 'kyActivateTerminal'  " & Err.Description
												kyActivateTerminal=false												
										End If
								Else

					End If 
		
		End Function
		
		
'=========================================================================================================================================================

		Function kyIfObjectExistsShortWait( sTestObjectName)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist(5)) Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												kyIfObjectExistsShortWait=true
										Else
												sKeywordError = "Error in 'kyIfObjectExistsShortWait'  " & Err.Description
												kyIfObjectExistsShortWait=false												
										End If
								Else
												sKeywordError = "Error in 'kyIfObjectExistsShortWait'. Object  " & CStr(sTestObjectName) & "  NOT Found."	
								End If
					End If 
		
		End Function

'=========================================================================================================================================================
' This Function select/click the menu item.


		Function kySelectFromMenu(menuNavigation)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								ExecuteString=getMenuExcuteString(menuNavigation)
								Execute ExecuteString &".Select"

								If Err.Number = 0 AND Err.Description = ""  Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kySelectFromMenu'  " & Err.Description	
								End If
								
					End If 
		
		End Function

'=========================================================================================================================================================

		Function kyWait(forSeconds)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								
								wait forSeconds
								
								If Err.Number = 0 AND Err.Description = ""  Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyWait'  " & Err.Description	
								End If
								
					End If 
		
		End Function



'=========================================================================================================================================================
' This Function Verifies if the Object is "Present on Page and Enabled"
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- None. But it sets the value of "bKeywordPF" to "False" if the Object is absent or Disabled.

		Function kyIfObjectNotExists(sTestObjectName)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist) Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyIfObjectNotExists'  " & Err.Description	
										End If
								Else
												sKeywordError = "Error in 'kyIfObjectNotExists'. Object  " & CStr(sTestObjectName) & "  NOT Found."	
								End If
					End If 
		
		End Function


'=========================================================================================================================================================
' This Function Verifies if the Object is "Present on Page and Enabled"
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- None. But it sets the value of "bKeywordPF" to "False" if the Object is absent or Disabled.

		Function kyIfObjectEnabled(sTestObjectName)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyIfObjectEnabled'  " & Err.Description	
										End If
								Else
												sKeywordError = "Error in 'kyIfObjectEnabled'. Object  " & CStr(sTestObjectName) & "  NOT Found on Page"	
								End If
					End If 
		
		End Function


'=========================================================================================================================================================
'This Function Verifies if the Object is "Present on Page and Disabled"
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- None. But it sets the value of "bKeywordPF" to "False" if the Object is absent or Enabled.

		Function kyIfObjectDisabled(sTestObjectName)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist AND objObject.GetROProperty("disabled") = 1) Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyIfObjectDisabled'  " & Err.Description	
										End If
								Else
												sKeywordError = "Error in 'kyIfObjectEnabled'. Object  " & CStr(sTestObjectName) & "  NOT Found OR NOT Disabled on Page"	
								End If
					End If 
		
		End Function


'=========================================================================================================================================================
'This Function returns the object property value
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- None. But it sets the value of "bKeywordPF" to "False" if the Object is absent or Enabled.

		Function kyGetObjectProperty( sTestObjectName, sObjProperty)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist) Then
									kyGetObjectProperty = objObject.GetROProperty(sObjProperty)
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyGetObjectProperty'  " & Err.Description	
										End If
								Else
												sKeywordError = "Error in 'kyGetObjectProperty'. Object  " & CStr(sTestObjectName) & "  NOT Found on Page"	
								End If
					End If 
		
		End Function
		
	
'=========================================================================================================================================================
	
		
		Function kySetObjectProperty( sTestObjectName, sObjProperty,val)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist) Then
										objObject.SetTOProperty sObjProperty, val
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												kySetObjectProperty=true
										Else
												sKeywordError = "Error in 'kySetObjectProperty'  " & Err.Description
												kySetObjectProperty=false												
										End If
								Else
												sKeywordError = "Error in 'kySetObjectProperty'. Object  " & CStr(sTestObjectName) & "  NOT Found on Page"	
								End If
					End If 
		
		End Function

'=========================================================================================================================================================
		
		Function kyVerifyObjectProperty(sTestObjectName, prop, value)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								If (objObject.Exist AND objObject.GetROProperty(prop) = value) Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												kyVerifyObjectProperty= objObject.GetROProperty(prop)
										Else
												sKeywordError = "Error in 'kyVerifyObjectProperty'  " & Err.Description	
												kyVerifyObjectProperty= objObject.GetROProperty(prop)
										End If
								Else
												sKeywordError = "Error in 'kyVerifyObjectProperty'. Object  " & CStr(sTestObjectName) & "  NOT Found on Page"	
								End If
					End If 
		
		End Function
'=========================================================================================================================================================
		
		Function kyVerifyTerminalRow(sTestObjectName, sObjProperty, rowData)
		
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties
								
								If (objObject.Exist) Then
										Dim rowTextFromApplication, rowTextFromTestData
										rowTextFromApplication = Trim(Replace(objObject.GetROProperty(sObjProperty)," ",""))
										rowTextFromTestData=Trim(Replace(rowData,"~",""))										
									
										If Err.Number = 0 AND Err.Description = "" AND  rowTextFromApplication=rowTextFromTestData Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												kyVerifyTerminalRow="FromApp: " & rowTextFromApplication  & "<br>FromData: " & rowTextFromTestData
										Else
												sKeywordError = "Error in 'kyVerifyTerminalRow'  " & Err.Description	
												kyVerifyTerminalRow="FromApp: " & rowTextFromApplication  & "<br>FromData: " & rowTextFromTestData
										End If
								Else
												sKeywordError = "Error in 'kyVerifyTerminalRow'. Object  " & CStr(sTestObjectName) & "  NOT Found on Page"	
								End If
					End If 
		
		End Function
		
'=========================================================================================================================================================

	' this keyword check the fields given in a form of row are present 
	
	Function kyVerifyTerminalRowOfFields(RowOfFieldObjects,morePagesObject,pagination,totalRowsInOnepage)  ' pagination,totalRowsInOnepage these two not used but may be required while updating keyword in future.
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Dim rowFound,morePagesAvailable,pageNo
								Dim arrReturn(2)  'this array can accomodate three values
								rowFound=False
								pageNo=1
								arrReturn(0)=1
								Dim refRowNo,refColNo,FoundOnThisPage
								arrOfObjects=split(RowOfFieldObjects,"~")								
								Do		
										
										Set objObject = kyGetObject (arrOfObjects(0)) ' this is reference object
										If (objObject.Exist(3)) Then
												 refRowNo=objObject.GetROProperty("start row")
												 refColNo=objObject.GetROProperty("start column")									 
												 For i= 1 To ubound(arrOfObjects)									 
													Set objObject1 = kyGetObject (arrOfObjects(i))
													objObject1.SetTOProperty "start row", refRowNo
													'objObject1.SetTOProperty "start column", (refColNo+cint(arrOfObjects(i-1)))
													If (objObject1.Exist(1)) Then						
														FoundOnThisPage=True
													Else																		
														FoundOnThisPage=False
														Exit for
													End if
												 Next
												 If(i=(ubound(arrOfObjects)+1) and (FoundOnThisPage=true or ubound(arrOfObjects)=0)) Then
												 	rowFound=true
												 	arrReturn(1)=refRowNo
												 	arrReturn(2)=refColNo
												 	objObject.SetTOProperty "start row", refRowNo ' here we set the row no of the found element as OR we have not added start row as property.
												 	Exit Do
												 End If
										Else
											Set objMore = kyGetObject (morePagesObject)
											morePagesAvailable=objMore.Exist(1)
											If morePagesAvailable=true Then
												Set WshShell = CreateObject("WScript.Shell")
												WshShell.SendKeys "{PGDN}"	
												pageNo=pageNo+1
												arrReturn(0)=pageNo	
											End if	
										End If

								Loop While morePagesAvailable=true
						
								If Err.Number = 0 AND Err.Description = "" AND  rowFound=true Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										kyVerifyTerminalRowOfFields=arrReturn
								Else
										sKeywordError = "Error in 'kyVerifyTerminalRowOfFields'  " & Err.Description	
										kyVerifyTerminalRowOfFields=false
								End If

					End If 
		
		End Function
		

'=========================================================================================================================================================

'This Function press the key.
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- None. But it sets the value of "bKeywordPF" to "False" if the Object is absent or Enabled.

		Function kyPress(keys)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								
								bKeywordPF = False
								Set WshShell = CreateObject("WScript.Shell")
								WshShell.SendKeys keys								
							
								If Err.Number = 0 AND Err.Description = ""  Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyPress'  " & Err.Description	
								End If								
					End If 
		
		End Function

'This Function press the key.
'Input:-  Object Type, Objects Properties and Values pairs separated by delimeter ","
'Output:- None. But it sets the value of "bKeywordPF" to "False" if the Object is absent or Enabled.

		Function kyPressMultipleTimes(keys, times)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
					bKeywordPF = False
					Set WshShell = CreateObject("WScript.Shell")
					
					For i = 1 To times								
								WshShell.SendKeys keys
					Next							
															
							
								If Err.Number = 0 AND Err.Description = ""  Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyPressMultipleTimes'  " & Err.Description	
								End If								
					End If 
		
		End Function

'=========================================================================================================================================================


		Function kyOpenApplication(path)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								
								bKeywordPF = False
								SystemUtil.Run path '							
							
								If Err.Number = 0 AND Err.Description = ""  Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyOpenApplication'  " & Err.Description	
								End If								
					End If 
		
		End Function
'=========================================================================================================================================================
				
		Function kyVerifyObjectWithPropertyExist( sTestObjectName,prop,val)
							On Error Resume Next		
							If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
										Dim objObject		
										bKeywordPF = False						
										Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties	
		
		
										objObject.SetTOProperty prop, empty
										objObject.SetTOProperty prop, val
		
										If objObject.Exist(5) Then
												bKeywordPF = True
												kyVerifyObjectWithPropertyExist=True
										Else
												kyVerifyObjectWithPropertyExist=false
												sKeywordError = "Error in 'kyVerifyObjectWithPropertyExist'.   "  & sTestObjectName & "  Not Found on Page  " & Err.Description	
										End If    								
							End If 
				
		End Function
		
'=========================================================================================================================================================
		

		Function kyVerifyObjectWithPropertyNotExist( sTestObjectName,prop,val)
							On Error Resume Next		
							If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
										Dim objObject		
										bKeywordPF = False						
										Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Properties	
		
		
										objObject.SetTOProperty prop, empty
										objObject.SetTOProperty prop, val
		
										If objObject.Exist(5)  = False Or objObject.Exist(5)  = Empty Then
												bKeywordPF = True
												kyVerifyObjectWithPropertyNotExist=True
										Else
												kyVerifyObjectWithPropertyNotExist=false
												sKeywordError = "Error in 'kyVerifyObjectWithPropertyNotExist'.   "  & sTestObjectName & "  Found on Page  " & Err.Description	
										End If    								
							End If 
				
		End Function
		
		
'=========================================================================================================================================================


'momin: this key update the genral object row and column with reference to the given refernce object.
Function kyUpdateRowColumnProperty(sGeneralTestObjectName, refRowNo, refColNo) ' here sTestObjectName is know and reference object and column and row no are tpo added or subtract to get the problem cordinate to write the no.
					On Error Resume Next
					If bKeywordPF = True then
								Dim objObject',refRowNo,refColNo
								bKeywordPF = False
'								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Name
'								If ( objObject.Exist(1)) Then 
'									refRowNo=objObject.GetROProperty("start row")
'									refColNo=objObject.GetROProperty("start column")
									Set objObject = kyGetObject (sGeneralTestObjectName)
'									If ( objObject1.Exist(1)) Then 
										objObject.SetTOProperty "start row", refRowNo
										objObject.SetTOProperty "start column", refColNo							
'									End if

									If Err.Number = 0 AND Err.Description = ""  Then
										
											bKeywordPF = true' Keyword Objective is Achived so value is Set to True
											kyUpdateRowColumnProperty=true
									Else
											sKeywordError = "Error in 'kygetObjectUsingReferenceObject'  " & Err.Description
											kyUpdateRowColumnProperty=false											
									End If
'								Else
'												sKeywordError = "Error in 'kygetObjectUsingReferenceObject'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
								'End If
					End If	 
		
		End Function

				

'=========================================================================================================================================================
' This Function display the message window if the object is not found.
Function kyGetDummyObject(sObjectClass)

			Dim objDummy
			Set objDummy = Description.Create
			objDummy("miccleass").Value = sObjectClass
			objDummy("Exist").Value = False
			objDummy("disabled").Value = 1
			'Execute "Set  kyGetDummyObject = Browser(sBrowser).Page(sPage)." & sObjectClass & "(objPropertyCollection)"
			Execute "Set  kyGetDummyObject = " & sObjectClass & "(objDummy)"
			
End Function



'=========================================================================================================================================================

Function getMenuExcuteString(menuNavigation)
				Dim menuNavigationTemp
				menuNavigationTemp=Replace(menuNavigation,JavaMenuName,"")
				menuItem=Split(menuNavigationTemp,">>")								
				For i = 0 to UBound(menuItem)								
					commonString=".JavaMenu(""XYZ"")"
					addString=Replace(commonString,"XYZ",menuItem(i))
					executeString= executeString+addString   
				Next
				
				getMenuExcuteString ="JavaWindow("""& win02 &""")" & executeString 
			
End Function





































'========================Function:  Table Cell Operation ========================

' This Function Performs the desired Operaion on the Web Object Present inside the Specific Table Cell
'Input:-  Objects Logical Name in the Shared Repository, Cell Text to find Row, Column no, Type of Object, Action, Value to be entered-selected if any
'Output:- None - Sets the value of "bKeywordPF" to True or False based on Operation performed.

Sub kyActionOnSpecificTableCell(sBrowserCreationTime,sTestObjectName, sCellText, iColumnNo, sMicClass, sAction, sText)

		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim objOTable, iRowNumber, objCellObject
						bKeywordPF = False
						Set objOTable = kyGetObject(sBrowserCreationTime, sTestObjectName) ' -- Get Object from Shared OR
						iRowNumber =  objOTable.GetRowWithCellText(sCellText)'Browser("Life Insurance Solution").Page("Life Insurance Solution").WebTable("column names:=" & sColumnNames).GetRowWithCellText(sCellText)
						Set objCellObject =  objOTable.ChildItem(iRowNumber, iColumnNo, sMicClass,0)' -- Fetch the Child Object from Specific Table Cell
						If ((objCellObject.Exist ="True")  AND ((objCellObject.GetROProperty("disabled") = 0) OR (objCellObject.GetROProperty("disabled") = "")))Then ' -- Check if the Object is Present and Enabled					
								if(sAction="Set" OR sAction="Select") Then
										Execute "objCellObject."  & sAction & " """ & sText & """"' -- Perform the Set or Select Action
								Else
										 Execute "objCellObject."  & sAction' -- Perform the Click Action
								End If
								
								If Err.Number = 0 AND Err.Description = ""  Then '-- Check if any Error Occured during Performing the above Operation
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyPerformOperation'  " & Err.Description	
								End If

						Else
								sKeywordError = "Error in 'kyActionOnSpecificTableCell'.   "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
						End If    ' End of -- (objCellObject.Exist AND objObject.GetROProperty("disabled") = 0) Then
		End If ' End of -- If bKeywordPF = True Then

End Sub



'========================Function:  Table Contains specified rows========================


		Function kyRowWithCellText(sObjectsProperties, sCellText, iColNo)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objPropertyCollection, objWebTable
								bKeywordPF = False
								Set objWebTable = kyGetObject ("WebTable", sObjectsProperties)' Get the Object with mentioned Description Properties
								If (objWebTable.Exist) Then
								kyRowWithCellText =  objWebTable.GetRowWithCellText(sCellText,iColNo,4)
										If kyRowWithCellText > -1 AND Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyRowWithCellText'  Row with Text " & sCellText & "NOT Found in Table " & Err.Description	
										End If
								Else
												sKeywordError = "Error in 'kyRowWithCellText'. Object  " & CStr(sObjectsProperties) & "  NOT Found on Page"	
								End If
					End If 
		
		End Function


' This Function determines whether the table contains specified rows 
'Input:-  Browser creation time,Objects Logical Name in the Shared Repository and NoofRows
'Output:- None - Sets the value of "bKeywordPF" to True or False based on row number count.

Function kyIfTableContainsSpecifiedRows(iBrowserCreationTime, sTestObjectName, iNoOfRows)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim objOTable, iRowNumber
						bKeywordPF = False
						Set objOTable = kyGetObject (iBrowserCreationTime, sTestObjectName) ' -- Get the Object 
						If objOTable.Exist Then
								iRowNumber =  objOTable.RowCount								
								If (iRowNumber = iNoOfRows) AND Err.Number = 0 AND Err.Description = "" Then ' -- if  row count matches and no error occured during Performing the above Operation							    
										bKeywordPF = True' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyIfTableContainsSpecifiedRows'  " & Err.Description	
								End If
						Else
								sKeywordError = "Error in 'kyIfTableContainsSpecifiedRows'.   "  & sTestObjectName & "  NOT Found on Page  " & Err.Description	
						End If    
		End If ' End of -- If bKeywordPF = True 

End Function


'========================Function:  Object Existence check for optional objects========================
' This Function Verifies if the Object is "Present".  It will not set the bKeywordPF to false even if the object not present on the page.
'Input:-  Browser creation time,Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ","
'Output:- None. 

Function kyIfObjectExistsShortWaitWeb( iBrowserCreationTime, sTestObjectName)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject								
								Set objObject = kyGetObject (iBrowserCreationTime, sTestObjectName)' Get the Object with mentioned Description Properties
								If objObject.Exist(40)  Then
										Exit Function
								End If
					End If 
		
End Function


'========================Function:  Object not exist========================
' This Function Verifies if the Object is not "Present" on the page.
'Input:-  Browser creation time,Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ","
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyIfObjectNOTOnPageShortWait(iBrowserCreationTime, sTestObjectName)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject		
								bKeywordPF = False						
								Set objObject = kyGetObject (iBrowserCreationTime, sTestObjectName)' Get the Object with mentioned Description Properties
								If objObject.Exist(5)  = False Or objObject.Exist(5)  = Empty Then
										bKeywordPF = True
								Else
										sKeywordError = "Error in 'kyIfObjectNOTOnPageShortWait'.   "  & sTestObjectName & "  Found on Page  " & Err.Description	
								End If    								
					End If 
		
End Function

'========================Function:  Retrieve data from Table Cell========================

' This Function retrieves the data from any  cell from the webtable
'Input:-  Browser creation time,Objects Logical Name in the Shared Repository, Cell Text to find Row, Column no
'Output:- The value retrieved from the cell .Also, sets the value of "bKeywordPF" to True or False based on Operation performed.

Function kyGetTableCellData(iBrowserCreationTime, sTestObjectName, sCellText, iColumnNo)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim objOTable, iRowNumber, sCellValue
						bKeywordPF = False
						Set objOTable = kyGetObject (iBrowserCreationTime, sTestObjectName) ' -- Get Object from Shared OR
						If objOTable.Exist  Then 						
								iRowNumber =  objOTable.GetRowWithCellText(sCellText)'Browser("Life Insurance Solution").Page("Life Insurance Solution").WebTable("column names:=" & sColumnNames).GetRowWithCellText(sCellText)
								sCellValue = objOTable.GetCellData(iRowNumber,iColumnNo)
								If  Err.Number = 0 AND Err.Description = "" Then
									bKeywordPF = True' Keyword Objective is Achived so value is Set to True
									kyGetTableCellData = sCellValue
								Else
										sKeywordError = "Error in 'kyGetTableCellData'  " & Err.Description	
								End If
						Else
								sKeywordError = "Error in 'kyGetTableCellData'.   "  & sTestObjectName & "  NOT Found on Page  " & Err.Description  						
						End If   
		End If ' End of -- If bKeywordPF = True Then

End Function

Function kyIPMGetTableCellData(iBrowserCreationTime, sTestObjectName, sCellText, iColumnNo,iRowNo)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim objOTable, iRowNumber, sCellValue
						bKeywordPF = False
						Set objOTable = kyGetObject (iBrowserCreationTime, sTestObjectName) ' -- Get Object from Shared OR
						If objOTable.Exist  Then 						
								iRowNumber =  objOTable.GetRowWithCellText(sCellText,,iRowNo)'Browser("Life Insurance Solution").Page("Life Insurance Solution").WebTable("column names:=" & sColumnNames).GetRowWithCellText(sCellText)
								sCellValue = objOTable.GetCellData(iRowNumber,iColumnNo)
								If  Err.Number = 0 AND Err.Description = "" Then
									bKeywordPF = True' Keyword Objective is Achived so value is Set to True
									kyIPMGetTableCellData = sCellValue
								Else
										sKeywordError = "Error in 'kyGetTableCellData'  " & Err.Description	
								End If
						Else
								sKeywordError = "Error in 'kyGetTableCellData'.   "  & sTestObjectName & "  NOT Found on Page  " & Err.Description  						
						End If   
		End If ' End of -- If bKeywordPF = True Then

End Function









'========================Function:  Compare the given value against the retrieved data from Table Cell========================

' This Function retrieves the data from any  cell from the webtable
'Input:-  Browser Creation Time,Objects Logical Name in the Shared Repository, Cell Text to find Row, Column no, Type of Object
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyCompareTableCellData(iBrowserCreationTime, sTestObjectName, sCellText, iColumnNo,stextToCompare)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim objOTable, iRowNumber, sCellValue
						bKeywordPF = False
						Set objOTable = kyGetObject (iBrowserCreationTime, sTestObjectName) ' -- Get Object from Shared OR
						If objOTable.Exist  Then 						
								iRowNumber =  objOTable.GetRowWithCellText(sCellText)'Browser("Life Insurance Solution").Page("Life Insurance Solution").WebTable("column names:=" & sColumnNames).GetRowWithCellText(sCellText)
								sCellValue = objOTable.GetCellData(iRowNumber,iColumnNo)
								If InStr(sCellValue,stextToCompare) > 0 AND Err.Number = 0 AND Err.Description = "" Then								
									bKeywordPF = True' Keyword Objective is Achived so value is Set to True									
								Else
										sKeywordError = "Error in 'kyCompareTableCellData'  " & Err.Description	
								End If
						Else
								sKeywordError = "Error in 'kyCompareTableCellData'.   "  & sTestObjectName & "  NOT Found on Page  " & Err.Description  						
						End If   
		End If ' End of -- If bKeywordPF = True Then

End Function

Function kyIPMCompareTableCellData(iBrowserCreationTime, sTestObjectName, sCellText, iColumnNo,iRowNo,stextToCompare)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim objOTable, iRowNumber, sCellValue
						bKeywordPF = False
						Set objOTable = kyGetObject (iBrowserCreationTime, sTestObjectName) ' -- Get Object from Shared OR
						If objOTable.Exist  Then 						
								iRowNumber =  objOTable.GetRowWithCellText(sCellText,,iRowNo)'Browser("Life Insurance Solution").Page("Life Insurance Solution").WebTable("column names:=" & sColumnNames).GetRowWithCellText(sCellText)
								sCellValue = objOTable.GetCellData(iRowNumber,iColumnNo)
								If InStr(sCellValue,stextToCompare) > 0 AND Err.Number = 0 AND Err.Description = "" Then								
									bKeywordPF = True' Keyword Objective is Achived so value is Set to True									
								Else
										sKeywordError = "Error in 'kyIPMCompareTableCellData'  " & Err.Description	
								End If
						Else
								sKeywordError = "Error in 'kyIPMCompareTableCellData'.   "  & sTestObjectName & "  NOT Found on Page  " & Err.Description  						
						End If   
		End If ' End of -- If bKeywordPF = True Then

End Function





'========================Function:  Wait till the object text changes========================

' This Function waits till the object text changes to the specified  value
'Input:-  Browser Creation Time, Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ","Previous Object Text and Expected text
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyWaitTillObjectTextChange(iBrowserCreationTime, sTestObjectName,sPreviousObjectText,sExpObjectText)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim bTextFlag,iCounter
						iCounter = 0
						'bKeywordPF = False
						bTextFlag = False
						Do
								iCounter = iCounter + 1
								If  iCounter < 20 Then								
										If  sExpObjectText = kyGetNextObjectText(iBrowserCreationTime, sTestObjectName,sPreviousObjectText) Then
											bTextFlag = True 
										Else
											Wait(5)
											Browser(Environment("sBrowser")).SetTOProperty "CreationTime", sBrowserCreationTime	
											Browser(Environment("sBrowser")).Page(Environment("sPage")).Refresh
										End If
								End If
						Loop Until bTextFlag = False	
						If  bTextFlag = True Then
							    If Err.Number = 0 AND Err.Description = "" Then
										bKeywordPF = True  'Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyWaitTillObjectTextChange'  " & Err.Description	
								End If
						Else
                                sKeywordError = "Error in 'kyWaitTillObjectTextChange' Object text not changed." 
						End If
		
		End If ' End of -- If bKeywordPF = True Then

End Function

'========================Function:  Click the object till it exists========================

' This Function clicks the object till it exists
'Input:-  Browser Creation Time, Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ","
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyClickObjTillItExsists(iBrowserCreationTime, sTestObjectName)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim bObjectFlag,objObject, iCounter
						iCounter = 0
						bKeywordPF = False
						bObjectFlag = True
						Do
								iCounter = iCounter + 1
								If  iCounter < 20 Then		
									objObject = kyIfObjectExistsShortWait(iBrowserCreationTime, sTestObjectName)
									If  objObject.Exist Then
										objObject.Click									
									Else
										bObjectFlag = False 
									End If
								End If
						Loop Until bObjectFlag = True	
						If  Err.Number = 0 AND Err.Description = "" Then
								bKeywordPF = True  'Keyword Objective is Achived so value is Set to True
					    Else
						 	   sKeywordError = "Error in 'kyClickObjTillItExsists'  " & Err.Description	
						End If
		End If ' End of -- If bKeywordPF = True Then

End Function

'========================Function:  Compare the property value against the expected========================

' This Function compares the property value against the expected value
'Input:-  Browser Creation Time, Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ",", Property  and Expected Value
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyCompareObjectProperty(iBrowserCreationTime, sTestObjectName,sObjProperty,sExpValue)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								'bKeywordPF = False
								If (sExpValue = Trim(kyGetObjectProperty( iBrowserCreationTime, sTestObjectName,sObjProperty))) Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyCompareObjectProperty'  " & Err.Description	
												bKeywordPF = False
										End If
								Else
										sKeywordError = "Error in 'kyCompareObjectProperty'. Property value doesn't match. "
										bKeywordPF = False
								End If
					End If 
		
End Function



' This Function compares the property value against the expected value
'Input:-  Browser Creation Time, Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ",", Property  and Expected Value
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyValueInString(sStringSearched,sSearchString)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								'bKeywordPF = False
								If (Instr(1,sStringSearched,sSearchString)<> 0)Then
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyValueInString'  " & Err.Description	
												bKeywordPF = False
										End If
								Else
										sKeywordError = "Error in 'kyValueInString'. String: "&sSearchString&" not found in String :"&sStringSearched
										bKeywordPF = False
								End If
					End If 
		
End Function







'========================Function:  Check if page exists========================

' This Function verifies that the correct page exists
'Input:-  Browser Creation Time and Expected page's title
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyIfPageExists( sExpPageTitle,iBrowserCreationTime)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								bKeywordPF = False	
								Browser(Environment("sBrowser")).SetTOProperty "CreationTime", iBrowserCreationTime
								Browser(Environment("sBrowser")).Page("title:="&sExpPageTitle).Sync							
								If (Browser(Environment("sBrowser")).Page("title:="&sExpPageTitle).Exist(40)) AND Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
												sKeywordError = "Error in 'kyIfPageExists' : Page "&sExpPageTitle&" :::" & Err.Description	
								End If
					End If 
		
End Function


'========================Function:  Navigate to the desired page========================

' This Function verifies that the correct page exists
'Input:-  Browser Creation Time and Objects Logical Name in the Shared Repository separated by ">". For Eg: Link_Home>Link_WorkFlow
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyNavigation( iBrowserCreationTime,sNavigationObjects)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								bKeywordPF = False	
								Dim iObjCount,objObject
								sNavigationObjects = Split(sNavigationObjects,">")	
                                For iObjCount = 0 to UBound(sNavigationObjects)
									Set objObject = kyGetObject (iBrowserCreationTime, sNavigationObjects(iObjCount))
									If objObject.Exist Then
												objObject.Click
									Else
												sKeywordError = "Error in 'kyNavigation'.   "  & objObject & "  NOT Found on Page  " & Err.Description  
												Exit For
									End If   
							   Next
							   If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
												sKeywordError = "Error in 'kyNavigation'  " & Err.Description	
								End If
					End If 
		
End Function
'=========================================================================================


' This Function verifies that the correct page exists
'Input:-  Browser Creation Time and Objects Logical Name in the Shared Repository , Property Name , Value to be set.
'Output:- None. - Sets the Object property  with the specifed  value ."bKeywordPF" to True or False based 

Function kySetORObjectProperty( iBrowserCreationTime,sMIcClass,sTestObjectName,sPropertyName,sValue)
					On Error Resume Next	
					Dim oORTestObject,sBrowser,sPage
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								bKeywordPF = False	


								sBrowser= Environment("sBrowser")
								sPage =Environment("sPage")

								Browser(sBrowser).SetTOProperty "CreationTime", iBrowserCreationTime	

						

								Set oORTestObject = Eval ("Browser(sBrowser).Page(sPage)."&sMicClass&"("""&sTestObjectName&""")")

								oORTestObject.SetTOProperty sPropertyName,sValue
					
							   If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
												sKeywordError = "Error in 'kyNavigation'  " & Err.Description	
								End If
					End If 
		
End Function




' ========== Keywords  End ========

' This Function verifies that the correct page exists
'Input:-  Browser Creation Time and Objects Logical Name in the Shared Repository , Property Name , Value to be set.
'Output:- None. - Sets the Object property  with the specifed  value ."bKeywordPF" to True or False based 

Function kyCloseBrowser (iBrowserCreationTime)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								bKeywordPF = False  
								Browser(Environment("sBrowser")).SetTOProperty "CreationTime", iBrowserCreationTime	
								If (Browser(Environment("sBrowser")).exist) Then 
										Browser(Environment("sBrowser")).close 
								Else
										sKeywordError = "Error in 'kyCloseBrowser'. Property value doesn't match. "
										bKeywordPF = False
								End If
					
								If Err.Number = 0 AND Err.Description = ""  Then
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyCloseBrowser'  " & Err.Description	
								End If
					End If 
		
End Function


'=================================================
' This Function Performs the desired Operation on the Web Object
'Input:-  Object Type, Objects Logical Name in the Shared Repository, Action to Perform, Value to be entered-selected if any
'Output:- Complete Test Object with Hirarchy from Broser

		Function kyGetItemWebList ( iBrowserCreationTime, sTestObjectName, iItemNumber)
					On Error Resume Next
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (iBrowserCreationTime,sTestObjectName)' Get the Object with mentioned Description Name
								If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty


											kyGetItemWebList = objObject.GetItem (iItemNumber)
	
											If Err.Number = 0 AND Err.Description = ""  Then
													bKeywordPF = true' Keyword Objective is Achived so value is Set to True
											Else
													sKeywordError = "Error in 'kyPerformOperation'  " & Err.Description	
											End If
								Else
												sKeywordError = "Error in 'kyPerformOperation'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
								End If
					End If	 
		
		End Function
'=================================================
 'This Function Performs the desired Operation on the Web Object
'Input:-  Object Type, Objects Logical Name in the Shared Repository, Action to Perform, Value to be entered-selected if any
'Output:- Complete Test Object with Hirarchy from Broser

		Function kyGetItemWebList ( iBrowserCreationTime, sTestObjectName, iItemNumber)
					On Error Resume Next
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = kyGetObject (iBrowserCreationTime,sTestObjectName)' Get the Object with mentioned Description Name
								If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty


											kyGetItemWebList = objObject.GetItem (iItemNumber)
	
											If Err.Number = 0 AND Err.Description = ""  Then
													bKeywordPF = true' Keyword Objective is Achived so value is Set to True
											Else
													sKeywordError = "Error in 'kyPerformOperation'  " & Err.Description	
											End If
								Else
												sKeywordError = "Error in 'kyPerformOperation'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
								End If
					End If	 
		
		End Function
'=================================================



	Function kyGetWinObjectProperty( sCompleteTestObjectName, sObjProperty)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								bKeywordPF = False
								Set objObject = sCompleteTestObjectName' Get the Object with mentioned Description Properties
								If (objObject.Exist) Then
									kyGetObjectProperty = objObject.GetROProperty(sObjProperty)
										If Err.Number = 0 AND Err.Description = ""  Then
												bKeywordPF = true' Keyword Objective is Achived so value is Set to True
										Else
												sKeywordError = "Error in 'kyIfObjectExists'  " & Err.Description	
										End If
								Else
												sKeywordError = "Error in 'kyIfObjectExists'. Object  " & CStr(sObjectsProperties) & "  NOT Found on Page"	
								End If
					End If 
		
		End Function
		
		
'This function will perform any action on any cell of web table.
'Input:-  Objects Logical Name in the Shared Repository, Column name, Cell text value for operation, Type of Object, Action, Value to be entered-selected if any
'Output:- Performs the action and Sets the value of "bKeywordPF" to True or False based on Operation performed.
'Note: Either sColName or sCellText value is necessary to pass.

Public Function kyActionOnAnyTableCell(sBrowserCreationTime,sTestObjectName, sColName, sCellText, sMicClass, sAction, sText)
		On Error Resume Next
		If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
						Dim objOTable, intRowNumber, intColCount, intRowCount, objCellObject, intColumnIndex, intCnt, intChildItemCnt, intRowData
						bKeywordPF = False
						Set objOTable = kyGetObject(sBrowserCreationTime, sTestObjectName) ' -- Get Object from Shared OR
						intRowCount =  objOTable.GetROProperty("rows")
						If sCellText = "" Then
								For i = 1 to intRowCount-1
									intColCount =  objOTable.ColumnCount(i)
									If intColCount >1 Then
										intRowData = i
										Exit For
									End If
								Next							
								For intColIteration = 1 to intColCount
									strColName = objOTable.GetCellData(1,intColIteration)
									If Trim(Ucase(sColName)) = Trim(Ucase(strColName)) Then
										intColumnIndex = intColIteration
										Exit For
									End If
								Next
								If intRowCount >intRowData Then
									For intCnt =  intRowData +1 to intRowCount
										intChildItemCnt = objOTable.ChildItemCount(intCnt, intColumnIndex, sMicClass)
										If intChildItemCnt>0 Then
											Set objCellObject =  objOTable.ChildItem(intCnt, intColumnIndex, sMicClass,intChildItemCnt-1)' -- Fetch the Child Object from Specific Table Cell
											err.clear
											Exit For
										End If
									Next
								Else
									Set objCellObject =  objOTable.ChildItem(intRowCount, intColumnIndex, sMicClass,0)' -- Fetch the Child Object from Specific Table Cell
								End If
						Else								
								For i = 1 to intRowCount-1
										intRowData =  objOTable.GetRowWithCellText(sCellText,,i)
										If intRowData > 0 Then
											intColCount =  objOTable.ColumnCount(intRowData)
											For j = 1 to intRowData
													For intColIteration = 1 to intColCount
															strColName = objOTable.GetCellData(j,intColIteration)
															If Trim(Ucase(sColName)) = Trim(Ucase(strColName)) Then
																intColumnIndex = intColIteration
																err.clear
																Exit For
															End If
													Next
													If intColumnIndex <> "" Then
														Exit for
													End If
											Next
											err.clear
										End If
										If intColumnIndex <> "" Then
											Exit for
										End If
								Next
								
'                                intColCount =  objOTable.ColumnCount(intRowData)
'								For i = 1 to intRowData
'										For intColIteration = 1 to intColCount
'												strColName = objOTable.GetCellData(i,intColIteration)
'												If Trim(Ucase(sColName)) = Trim(Ucase(strColName)) Then
'													intColumnIndex = intColIteration
'													err.clear
'													Exit For
'												End If
'										Next
'										If intColumnIndex <> "" Then
'											Exit for
'										End If
'								Next
								
								Set objCellObject =  objOTable.ChildItem(intRowData, intColumnIndex, sMicClass,0)' -- Fetch the Child Object from Specific Table Cell
						End If
						
                        If ((objCellObject.Exist ="True")  AND ((objCellObject.GetROProperty("disabled") = 0) OR (objCellObject.GetROProperty("disabled") = "")))Then ' -- Check if the Object is Present and Enabled					
								if(sAction="Set" OR sAction="Select") Then
										Execute "objCellObject."  & sAction & " """ & sText & """"' -- Perform the Set or Select Action
                                Elseif Instr(UCase(sAction),"GETROPROPERTY") > 0 Then																																		
										err.clear
										strPoprName= Split(sAction,"_")(1)
										kyActionOnAnyTableCell = Trim(objCellObject.GetROProperty(strPoprName))
								Else
										 Execute "objCellObject."  & sAction' -- Perform the Click Action
								End If
								
								If Err.Number = 0 AND Err.Description = ""  Then '-- Check if any Error Occured during Performing the above Operation
										bKeywordPF = true' Keyword Objective is Achived so value is Set to True
								Else
										sKeywordError = "Error in 'kyActionOnAnyTableCell'  " & Err.Description	
								End If

						Else
								sKeywordError = "Error in 'kyActionOnAnyTableCell'.   "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
						End If    ' End of -- (objCellObject.Exist AND objObject.GetROProperty("disabled") = 0) Then
		End If ' End of -- If bKeywordPF = True Then
		On Error goto 0

End Function


'=================================================
 'This Function Performs the desired Operation on the Win Object
 'Input:-  Object Type, sMicClass,Objects Logical Name in the Shared Repository, Action to Perform, Value to be entered-selected if any
'Output:-Action performed on the Required object and bKeywordPF set accordingly.



'Function kyPerformWinOperation (sCompleteTestObjectName, sAction, sText)
'					On Error Resume Next
'					If bKeywordPF = True AND ( sAction = "Click" OR  sText <> Empty) Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
'								Dim objObject
'								bKeywordPF = False
'								Set objObject = sCompleteTestObjectName' Get the Object with mentioned Description Name
'								If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty
'											if( sAction = "Click" ) Then
'													Execute "objObject." & sAction'objObject.Click
'											ElseIf (UCase(sText) = "BLANK") Then' Only Applicable for WebEdit
'													Execute "objObject." & sAction & " """""
'											Else
'													Execute "objObject." & sAction & " """ & sText & """"     ' ==  Split(sText, cDATADELIMETER)(iFunctionIteration - 1) & """"
'											End If
'	
'											If Err.Number = 0 AND Err.Description = ""  Then
'													bKeywordPF = true' Keyword Objective is Achived so value is Set to True
'											Else
'													sKeywordError = "Error in 'kyPerformWinOperation'  " & Err.Description	
'											End If
'								Else
'												sKeywordError = "Error in 'kyPerformWinOperation'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
'								End If
'					End If	 
'		
'		End Function
'
'
'



		'==============================================
' This Function Performs the desired Operation on the Web Object if object Exist but doesnot impact the test status
'Input:-  Object Type, sMicClass,Objects Logical Name in the Shared Repository, Action to Perform, Value to be entered-selected if any
'Output:- if the Object exist then the required action will be performed but the bKeywordPF is not set hence this keyword will not impact the test case status.

		Function kyPerformWebOperationIfExist ( iBrowserCreationTime, sMicClass,sTestObjectName, sAction, sText)
					On Error Resume Next
					If bKeywordPF = True AND ( sAction = "Click" OR  sText <> Empty) Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim objObject
								'bKeywordPF = False
								sBrowser= Environment("sBrowser")
								sPage =Environment("sPage")

								Set objObject =Eval ("Browser(sBrowser).Page(sPage)."&sMicClass&"("""&sTestObjectName&""")")' Set the Required Object.
								If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty
											if( sAction = "Click" ) Then
													Execute "objObject." & sAction'objObject.Click
											ElseIf (UCase(sText) = "BLANK") Then' Only Applicable for WebEdit
													Execute "objObject." & sAction & " """""
											Else
													Execute "objObject." & sAction & " """ & sText & """"     ' ==  Split(sText, cDATADELIMETER)(iFunctionIteration - 1) & """"
											End If
	
'											If Err.Number = 0 AND Err.Description = ""  Then
'													bKeywordPF = true' Keyword Objective is Achived so value is Set to True
'											Else
'													sKeywordError = "Error in 'kyPerformOperation'  " & Err.Description	
'											End If
'								Else
'												sKeywordError = "Error in 'kyPerformOperation'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
								End If
					End If	 
		
		End Function
'=================================================

' This Function Performs the desired Operation on the child object of a particular mic type.
'Input:-  strObjectClass - Mic Class,strObjectProperty- Property Name,strObjectValue- Property value ,strAction -,strOptionalText ,intOptionalChilditemNumber- child item number on which the action is to be performed.
'Output:- if the Object exist then the required action will be performed and the bKeywordPF  will be set accordingly.
Function kyAnyActionOnChildItem(strObjectClass,strObjectProperty,strObjectValue,strAction,strOptionalText,intOptionalChilditemNumber)

		On Error Resume Next
		If bKeywordPF = True Then
					bKeywordPF = False
					Set objObject = Description.Create( )
					objObject("micclass").Value = strObjectClass
					objObject(strObjectProperty).Value = strObjectValue

					Set objChild =Browser(Environment("sBrowser")).Page(Environment("sPage")).ChildObjects(objObject)
'					Set objChild =kyGetChildObject (iBrowserCreationTime, strObjectClass, strObjectProperty, strObjectValue)
					intItemCnt = objChild.count

					If  intItemCnt > 0 Then            

								If intOptionalChilditemNumber = "" Or Ucase(intOptionalChilditemNumber) = "FIRST" Then
												intOptionalChilditemNumber = 0
								ElseIf Ucase(intOptionalChilditemNumber) = "LAST" Then
												intOptionalChilditemNumber = intItemCnt-1
								End If

								If ( objChild(intOptionalChilditemNumber).Exist AND objChild(intOptionalChilditemNumber).GetROProperty("disabled") = 0) Then
											if( strAction = "Click" ) Then
															Execute "objChild(intOptionalChilditemNumber)." & strAction'objObject.Click
											ElseIf (Trim(UCase(strOptionalText)) = "BLANK") Then' Only Applicable for WebEdit
															Execute "objChild(intOptionalChilditemNumber)." & strAction & " """""
											Else
															Execute "objChild(intOptionalChilditemNumber)." & strAction & " """ & strOptionalText & """"     ' ==  Split(sText, cDATADELIMETER)(iFunctionIteration - 1) & """"
											End If

											If Err.Number = 0 AND Err.Description = ""  Then
															bKeywordPF = true' Keyword Objective is Achived so value is Set to True
											Else
															sKeywordError = "Error in 'kyAnyActionOnChildItem'  " & Err.Description  
											End If
								Else
											sKeywordError = "Error in 'kyAnyActionOnChildItem'. Object NOT Found OR NOT Enabled on Page  " & Err.Description                
								End If
					Else
								sKeywordError = "Error in 'kyAnyActionOnChildItem'.  Object Not Found  " & Err.Description    
					End If
		End If
                
End Function

Function kyAnyActionOnChildItemMultiBrowser(iBrowserCreationTime, strObjectClass,strObjectProperty,strObjectValue,strAction,strOptionalText,intOptionalChilditemNumber)

		On Error Resume Next
		If bKeywordPF = True Then
					bKeywordPF = False

					Set objChild =kyGetChildObject (iBrowserCreationTime, strObjectClass, strObjectProperty, strObjectValue)
					intItemCnt = objChild.count

					If  intItemCnt > 0 Then            

								If intOptionalChilditemNumber = "" Or Ucase(intOptionalChilditemNumber) = "FIRST" Then
												intOptionalChilditemNumber = 0
								ElseIf Ucase(intOptionalChilditemNumber) = "LAST" Then
												intOptionalChilditemNumber = intItemCnt-1
								End If

								If ( objChild(intOptionalChilditemNumber).Exist AND objChild(intOptionalChilditemNumber).GetROProperty("disabled") = 0) Then
											if( strAction = "Click" ) Then
															Execute "objChild(intOptionalChilditemNumber)." & strAction'objObject.Click
											ElseIf (Trim(UCase(strOptionalText)) = "BLANK") Then' Only Applicable for WebEdit
															Execute "objChild(intOptionalChilditemNumber)." & strAction & " """""
											Else
															Execute "objChild(intOptionalChilditemNumber)." & strAction & " """ & strOptionalText & """"
											End If

											If Err.Number = 0 AND Err.Description = ""  Then
															bKeywordPF = true' Keyword Objective is Achived so value is Set to True
											Else
															sKeywordError = "Error in 'kyAnyActionOnChildItemMultiBrowser'  " & Err.Description  
											End If
								Else
											sKeywordError = "Error in 'kyAnyActionOnChildItemMultiBrowser'. Object NOT Found OR NOT Enabled on Page  " & Err.Description                
								End If
					Else
								sKeywordError = "Error in 'kyAnyActionOnChildItemMultiBrowser'.  Object Not Found  " & Err.Description    
					End If
		End If
                
End Function

Function kyGetChildObject (iBrowserCreationTime, strObjectClass, strObjectProperty, strObjectValue)

			On Error Resume Next
			iBrowserCreationTimeGlobal = sBrowserCreationTime
			Dim sBrowserName, sPageTitle, objAppObject, sObjectClass, objPropertyCollection, sPropertyValuePairs
			Browser(Environment("sBrowser")).SetTOProperty "CreationTime",  iBrowserCreationTime
            sBrowser = Environment("sBrowser")
			sPage =Environment("sPage")
            
			Set objObject = Description.Create( )
			objObject("micclass").Value = strObjectClass
			objObject(strObjectProperty).Value = strObjectValue

			Set objChild =Browser(Environment("sBrowser")).Page(Environment("sPage")).ChildObjects(objObject)

			If objChild.Count < 1 Then
					bKeywordPF = False
					sKeywordError = "Error in 'kyGetChildObject'. Object NOT Found OR NOT Enabled on Page"
			Else
					Set kyGetChildObject = objChild				
			End If

End Function
'=================================================
' This Function checks the sorting on given string
'Input:-  strValue - String value contains different element seperated by semi-column (;)
'Output:- if the Object exist then the required action will be performed and the bKeywordPF  will be set accordingly.
Function kyCheckSorting(strValue)

		On Error Resume Next
		If bKeywordPF = True Then
                Dim arrSplitString, intCount, flgReturn
                flgReturn = "True"     
                arrSplitString = Split(strValue,";")
                
                For intCount = Lbound(arrSplitString) to UBound(arrSplitString) -1
						If arrSplitString(intCount) > arrSplitString(intCount +1) Then
								flgReturn = "False"
                                Exit For
						End If
                Next
                If flgReturn = "True" Then
						bKeywordPF = True' Keyword Objective is Achived so value is Set to True
				Else
						bKeywordPF = False
						sKeywordError = "Error in 'kyCheckSorting'  " & Err.Description	
				End If
		End If   

End Function

'=================================================

' This Function checks that whether the childitem of the given object is exist on page or not.
'Input:-  strObjectClass - Mic Class,strObjectProperty- Property Name,strObjectValue- Property value ,strExist -Null, True or False
'Output:- if the Object exist then the required action will be performed and the bKeywordPF  will be set accordingly.
Function kyIfChildItemExist(strObjectClass,strObjectProperty,strObjectValue,strExist)

		On Error Resume Next
		If bKeywordPF = True Then
					Set objObject = Description.Create( )
					objObject("micclass").Value = strObjectClass
					objObject(strObjectProperty).Value = strObjectValue

					Set objChild =Browser(Environment("sBrowser")).Page(Environment("sPage")).ChildObjects(objObject)
					intItemCnt = objChild.count

					If  intItemCnt > 0 And Err.Number = 0 Then
						If Trim(Ucase(strExist)) = "TRUE" Or Trim(Ucase(strExist)) = "" Then
							bKeywordPF = true' Keyword Objective is Achived so value is Set to True
						Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyIfChildItemExist'. Object SHOULD NOT exist on Page  " & Err.Description
						End If						
					Elseif Trim(Ucase(strExist)) = "FALSE" Then
						bKeywordPF = true' Keyword Objective is Achived so value is Set to True
					Else
						bKeywordPF = false
						sKeywordError = "Error in 'kyIfChildItemExist'. Object NOT Found OR Check the function parameters  " & Err.Description                
					End If
								
		End If
		On Error Goto 0
                
End Function

'=================================================

' This Function access the ecxel sheet and checks the column sequence.
'Input:-  strExcelPath - Path of excel file,strColName- Column name sequence with ;$ delimeter,strSheetName- Name of sheet in which column sequence need to be checked. 
'Output:- bKeywordPF  will be set accordingly.
Function kyExcelSheetAccess(strExcelPath,strColName,strSheetName)

		On Error Resume Next
		If bKeywordPF = True Then
					arrColName = split(strColName, ";$")					
					If strExcelPath = "" Then
						Set objExcel = GetObject ("" ,"Excel.Application")					
						Set objWB = objExcel.WorkBooks(1)
					Else
						Set objExcel = CreateObject ("Excel.Application")
	                    Set objWB = objExcel.WorkBooks.Open(strExcelPath)
					End If
					If Err.Number = 0 Then
							Set objWS = objWB.WorkSheets(strSheetName)
							Dim i
							i =1
							Do 
								strStd = objWS1.Cells(1, i).Value
								If strStd <> arrColName(i-1) Then
									sKeywordError = "Error in 'kyExcelSheetAccess'. Column not found or check the sequence."
									Exit do
								End If
								i = i + 1
							Loop while (objWS1.Cells(1, i).Value <>"")
							
							If i <> UBound(arrColName)+1 Then
								sKeywordError = "Error in 'kyExcelSheetAccess'. Column not found or check the sequence."
							Else
								bKeywordPF = true
							End If
					Else
								sKeywordError = "Error in 'kyExcelSheetAccess'. Error in excel object" & Err.Description
					End If
		End If
		On Error goto 0

End Function
					
'========================Function:  Checks the tool tip value of any object========================

' This Function checks the tool tip value of any object
'Input:-  Browser Creation Time, Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ",", Property  and Expected Value, Whether it should be present or not
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyCheckToolTipValue(iBrowserCreationTime, sTestObjectName,sObjProperty,sExpValue,sExist)
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
					If sObjProperty = "" Then
							strRunTimeValue = Trim(kyGetObjectProperty( iBrowserCreationTime, sTestObjectName,"outerhtml"))
					Else
							strRunTimeValue = Trim(kyGetObjectProperty( iBrowserCreationTime, sTestObjectName,sObjProperty))
					End If
					If strRunTimeValue<>"" AND sExpValue<>"" Then
							strRunTimeValue = UCase(Replace(strRunTimeValue," ",""))
							sExpValue = UCase(Replace(sExpValue," ",""))
							If InStr(strRunTimeValue, sExpValue) > 0 Then
									If sExist = "" OR Ucase(sExist) = "TRUE" Then
											If Err.Number = 0 AND Err.Description = ""  Then
													bKeywordPF = true' Keyword Objective is Achived so value is Set to True
											Else
													sKeywordError = "Error in 'kyCheckToolTipValue'  " & Err.Description	
													bKeywordPF = False
											End If
									Else
											sKeywordError = "Error in 'kyCheckToolTipValue'. Expected value should not matched with actual value"
											bKeywordPF = False
									End If												
							ElseIf Ucase(sExist) = "FALSE" Then
									If Err.Number = 0 AND Err.Description = ""  Then
											bKeywordPF = true' Keyword Objective is Achived so value is Set to True
									Else
											sKeywordError = "Error in 'kyCheckToolTipValue'  " & Err.Description	
											bKeywordPF = False
									End If
							Else
									sKeywordError = "Error in 'kyCheckToolTipValue'. Expected value should matched with actual value"
									bKeywordPF = False
							End If																		
					Else
							sKeywordError = "Error in 'kyCheckToolTipValue'. Expected or Run Time value is not present. "
							bKeywordPF = False
					End If
		End If
		On Error GoTo 0
		
End Function

'========================Function:  Table Cell Operation ========================

' This Function Performs the desired Operation on the Web Object Present inside the Specific Table Cell
'Input:-  Objects Logical Name in the Shared Repository, Cell Text to find Row, Column no, Type of Object, Action, Value to be entered-selected if any'Output:- None - Sets the value of "bKeywordPF" to True or False based on Operation performed.

Sub kyIPMActionOnSpecificTableCell(sBrowserCreationTime,sTestObjectName, sCellText, iColumnNo,iRowNo,sMicClass, sAction, sText)

	If bKeywordPF = True Then ' -- Check if the Previous Keyword was executed Successfully
			Dim objOTable, iRowNumber, objCellObject
			bKeywordPF = False
			Set objOTable = kyGetObject(sBrowserCreationTime, sTestObjectName) ' -- Get Object from Shared OR	
			iRowNumber =  objOTable.GetRowWithCellText(sCellText,,iRowNo)'Browser("Life Insurance Solution").Page("Life Insurance Solution").WebTable("column names:=" & sColumnNames).GetRowWithCellText(sCellText)
			Set objCellObject =  objOTable.ChildItem(iRowNumber, iColumnNo, sMicClass,0)
			If sMicClass = "WebRadioGroup" Then
					strAllItems = objCellObject.getROProperty("all items")
					sText = Split(strAllItems,";")(iRowNumber-iRowNo)
			End If                                                               
			If ((objCellObject.Exist ="True")  AND ((objCellObject.GetROProperty("disabled") = 0) OR (objCellObject.GetROProperty("disabled") = "")))Then ' -- Check if the Object is Present and Enabled                                                   
					if(sAction="Set" OR sAction="Select") Then
							Execute "objCellObject."  & sAction & " """ & sText & """"' -- Perform the Set or Select Action
					Else
							Execute "objCellObject."  & sAction' -- Perform the Click Action
					End If
					
					If Err.Number = 0 AND Err.Description = ""  Then '-- Check if any Error Occured during Performing the above Operation
							bKeywordPF = true' Keyword Objective is Achived so value is Set to True
					Else
							sKeywordError = "Error in 'kyPerformOperation'  " & Err.Description           
					End If
	
			Else
					sKeywordError = "Error in 'kyIPMActionOnSpecificTableCell'.   "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description     
			End If  
	End If 

End Sub



' This Function Performs two clicks  on a Object
'Input:-  Object Type, Objects Logical Name in the Shared Repository, 
'Output:- if the Object exist then the required action will be performed but the bKeywordPF is not set hence this keyword will not impact the test case status.
Function kyPerformTwoClicks ( iBrowserCreationTime,sTestObjectName)
		On Error Resume Next
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								
				Dim objObject
				bKeywordPF = False
				Set objObject = kyGetObject (iBrowserCreationTime,sTestObjectName)' Get the Object with mentioned Description Name
				If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty
							Dim QTPApp 
							Set QTPApp = CreateObject("QuickTest.Application")
							QTPApp.Test.Settings.Web.BrowserNavigationTimeout = 0
							objObject.click         
							objObject.click         
							QTPApp.Test.Settings.Web.BrowserNavigationTimeout = 60000               
							Set App = Nothing

							If Err.Number = 0 AND Err.Description = ""  Then
									bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							Else
									sKeywordError = "Error in 'kyPerformTwoClicks'  " & Err.Description     
							End If
				Else
							sKeywordError = "Error in 'kyPerformTwoClicks'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description 
				End If
								
		End If   

End Function



' This Function Performs Mouse clicks (Right or Left  on a Object
'Input:-  Object Type, Objects Logical Name in the Shared Repository, Mosue click Left or right
'Output:- if the Object exist then the required action will be performed but the bKeywordPF is not set hence this keyword will not impact the test case status.
Function kyPerformMouseClick( iBrowserCreationTime,sTestObjectName,sButton)
	On Error Resume Next
	If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
							
			Dim objObject
			bKeywordPF = False
			Set objObject = kyGetObject (iBrowserCreationTime,sTestObjectName)' Get the Object with mentioned Description Name
			If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty
					Setting.WebPackage("ReplayType") = 2 '
					If UCase(Trim(sButton))="RIGHT" Then
							objObject.FireEvent "Onclick",,,micRightBtn                                                                                                                                     
					Else
							objObject.FireEvent "onmouseover",,,0
							wait(1)
							objObject.FireEvent "Onclick",,,micLeftBtn 
					End If

					wait(1)
					Setting.WebPackage("ReplayType") = 1 '                                                                          

					If Err.Number = 0 AND Err.Description = ""  Then
							bKeywordPF = true' Keyword Objective is Achived so value is Set to True
					Else
							sKeywordError = "Error in 'kyPerformMouseClick'  " & Err.Description    
					End If
			Else
					sKeywordError = "Error in 'kyPerformMouseClick'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description        
			End If
							
	End If   

End Function


'#############################################################################################################

Function kyChildItemProperty(strMicClass,strObjectPropertyName,strObjectPropertyValue,strRequiredPropertyValue)

		On Error Resume Next
		If bKeywordPF = True Then
					bKeywordPF = False
					Set objChild =kyGetChildObject (0, strMicClass, strObjectPropertyName, strObjectPropertyValue)
					intItemCnt = objChild.count

					If  intItemCnt > 0 And Err.Number = 0 Then                                                                       
							bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyChildItemProperty = objChild(intItemCnt-1).GetROProperty(strRequiredPropertyValue)
					Else                                                               
							sKeywordError = "Error in 'kyChildItemProperty'. Object NOT Found OR Check the function parameters  " & Err.Description                
					End If                                                                                                
		End If

		On Error Goto 0
                
End Function

Function kyChildItemMouseClick( iBrowserCreationTime,strMicClass,strObjectPropertyName,strObjectPropertyValue,sButton,iChildItemNumber)
	On Error Resume Next
	If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action							
			Dim objObject
			Set objObject = kyGetChildObject (iBrowserCreationTime, strMicClass, strObjectPropertyName, strObjectPropertyValue) 'Get the Object with mentioned Description Name
			intItemCnt = objObject.count

			If  intItemCnt > 0 Then
					If iChildItemNumber = "" Or Ucase(iChildItemNumber) = "FIRST" Then
							iChildItemNumber = 0
					ElseIf Ucase(iChildItemNumber) = "LAST" Then
							iChildItemNumber = intItemCnt-1
					End If
	
					If ( objObject(iChildItemNumber).Exist AND objObject(iChildItemNumber).GetROProperty("disabled") = 0) Then
							Setting.WebPackage("ReplayType") = 2 '
							If UCase(Trim(sButton))="RIGHT" Then
									objObject(iChildItemNumber).FireEvent "Onclick",,,micRightBtn                                                                                                                                     
							Else
									objObject(iChildItemNumber).FireEvent "onmouseover",,,0
									wait(1)
									objObject(iChildItemNumber).FireEvent "Onclick",,,micLeftBtn 
							End If
	
							wait(1)
							Setting.WebPackage("ReplayType") = 1 '                                                                          
	
							If Err.Number = 0 AND Err.Description = ""  Then
											bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							Else
											sKeywordError = "Error in 'kyChildItemMouseClick'  " & Err.Description    
							End If
					Else
								sKeywordError = "Error in 'kyChildItemMouseClick'. "  & objObject & "  NOT Found OR NOT Enabled on Page  " & Err.Description        
					End If
			Else
					sKeywordError = "Error in 'kyChildItemMouseClick'. "  & objObject & "  NOT Found on Page"
			End If									
	End If   

End Function


'This function is to fire a Event on the same.
'Input:-  Objects Logical Name in the Shared Repository, Event , Button is the extra parameter for future use
'Output:- Performs the action and Sets the value of "bKeywordPF" to True or False based on Operation performed.



     Function kyFireEvent( iBrowserCreationTime,sTestObjectName,sEvent,sButton)
                                        On Error Resume Next
                                        If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
                                                                
                                                        Dim objObject
                                                                bKeywordPF = False
                                                                Set objObject = kyGetObject (iBrowserCreationTime,sTestObjectName)' Get the Object with mentioned Description Name
                                                                If ( objObject.Exist ) Then ' objObject <> Empty
                                                                                                        Setting.WebPackage("ReplayType") = 2 '
'                                                                                                                 
                                                                                                              objObject.FireEvent sEvent        
                                                                                                                                                                                                                                                                                                                                                                                                                                                  Setting.WebPackage("ReplayType") = 1'                                                         
                                        
                                                                                        If Err.Number = 0 AND Err.Description = ""  Then
                                                                                                        bKeywordPF = true' Keyword Objective is Achived so value is Set to True
                                                                                        Else
                                                                                                        sKeywordError = "Error in 'kyPerformMouseClick'  " & Err.Description    
                                                                                        End If
                                                                Else
                                                                                                sKeywordError = "Error in 'kyPerformMouseClick'. "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description        
                                                                End If
                                                                
                                        End If   
                End Function

' This Function compares the property value against the expected value with option 
'Input:-  Browser Creation Time, Objects Logical Name in the Shared Repository/Objects Properties and Values pairs separated by delimeter ",", Property  and Expected Value
'Output:- None. - Sets the value of "bKeywordPF" to True or False based on object existence.

Function kyValueInStringOptional(sStringSearched,sSearchString,strFlag)
					On Error Resume Next		
					If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
						
										If (Instr(1,sStringSearched,sSearchString) > 0) AND Err.Number = 0 AND Err.Description = "" Then											
												If Trim(strFlag) = "" or Trim(Ucase(Cstr(strFlag))) = "TRUE"  Then
														bKeywordPF = true' Keyword Objective is Achived so value is Set to True
												Else
														sKeywordError = "Error in 'kyCompareObjectProperty'. String: "&sSearchString&" not found in String :"&sStringSearched
														bKeywordPF = False
												End if

										Elseif Err.Number = 0 AND Err.Description = "" And Trim(Ucase(Cstr(strFlag))) = "FALSE"   Then												 
														bKeywordPF = true

										Else
													sKeywordError = "Error in 'kyValueInStringOptional'. String: "&sSearchString&" not found in String :"&sStringSearched
													bKeywordPF = False
												
										End If
					End If 
							
End Function



		
		
		
		
		
		
		
		'========================Function:  Table Cell Operation ========================

Function kyPerformWebOperation_____NEW ( sTestObjectName, sAction, sText)
					On Error Resume Next
					If bKeywordPF = True AND ( sAction = "Click" OR sText <> Empty) Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
								Dim  objObject
								bKeywordPF = False
								Set objObject = kyGetObject (sTestObjectName)' Get the Object with mentioned Description Name
								If ( objObject.Exist AND objObject.GetROProperty("disabled") = 0) Then ' objObject <> Empty
											if(sAction = "Click" ) Then
													Execute "objObject." & sAction'objObject.Click
											ElseIf (UCase(sText) = "BLANK") Then' Only Applicable for WebEdit
													Execute "objObject." & sAction & " """""
											Else
													Execute "objObject." & sAction & " """ & sText & """"
											End If
											' Check if any Error Occured during Performing the above Operation
											If Err.Number = 0 AND Err.Description = ""  Then
													bKeywordPF = true' Keyword Objective is Achived so value is Set to True
											Else
													sKeywordError = "Error in 'kyPerformOperation'  " & Err.Description	
											End If
								Else
												sKeywordError = "Error in 'kyPerformOperation'.   "  & sTestObjectName & "  NOT Found OR NOT Enabled on Page  " & Err.Description	
								End If
					End If	 
		
		End Function

'=========================================================================================================================================================
' This Function returns the Property Collection Object
'Input:-  Objects Properties and Values pairs separated by delimeter ","
'Output:- Property Collection Object

Function kyGetPropertiesCollection (sTestObjectName)

		Dim objPropColl, arrProperties, arrValues(), iPropertiesCount, arrObjectsPropertiesAndValues
		arrObjectsPropertiesAndValues = Split(sTestObjectName,",")
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

