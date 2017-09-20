'=========================================================================
'	TEST CASE TEMPLATE
'=========================================================================


'This function perform testorm XYZ.
Function TC_Template_Case (EXEC, UNIVERSAL)    
		
'		Sub_Login  EXEC, UNIVERSAL
'		Sub_AddGeneralCode  EXEC, UNIVERSAL
'		Sub_CloseApplication EXEC, UNIVERSAL


End Function




'=========================================================================
'	SUB TEMPLATE
'=========================================================================

'This sub perform XYZ operation
Sub Sub_Template (EXEC, UNIVERSAL)

	On Error Resume Next
	 If bMethodPF = True Then

				subInitializeFunctionVariables()
				UNIVERSAL.Item("TESTFLOW") = UNIVERSAL.Item("TESTFLOW") &">> Sub_Template "

				'*******************************************************************************************
				'write your code here
				
				'********************************************************************************************
				subMethodVerification EXEC, "Sub_Template", "Method Message"

	 End If

End Sub



'=========================================================================
'	KEY TEMPLATE
'=========================================================================

' This Key perform xyz operation
Function kyTemplate( iBrowserCreationTime, sObjectsProperties)
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
					Dim objObject
					bKeywordPF = False
					Set objObject = kyGetObject (iBrowserCreationTime, sObjectsProperties)' Get the Object with mentioned Description Properties
					If ((not isEmpty(objObject))AND objObject.GetROProperty("disabled") = 0) Then
							If Err.Number = 0 AND Err.Description = ""  Then
									bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							Else
									sKeywordError = "Error in 'kyTemplate'  " & Err.Description	
							End If
					Else
									sKeywordError = "Error in 'kyTemplate'. Object  " & CStr(sObjectsProperties) & "  NOT Found on Page"	
					End If
		End If 

End Function
