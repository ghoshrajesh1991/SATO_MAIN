	Sub errHandler(sObj)
	
				Dim sDetailedError
				Environment("ErrorMsg") = ""
'				Environment("ErrorNum") = ""
'				If Err.Number <> 0 Then
					Environment("ErrorNum") = Err.Number
'				End If
				
					arrObject = Split(sObj, ".")
					
					Select Case Environment("ErrorNum")
						 Case -2147220990, 424
						 	Environment("ErrorMsg") = "<b>"&arrObject(Ubound(arrObject)) &"</b>:Object not found on page.<br>"&_
						 							  "Please check the object properties.<br>"			
						 Case 13
							Environment("ErrorMsg") = "<b>"&arrObject(Ubound(arrObject)) &"</b>:Type mismatch error found while trying to perform operation on this object<br>"						 
							
						Case -2147220983
							Environment("ErrorMsg") = "The environment parameter not initialised"
						 							  	
					End Select
					
							If Environment("ErrorNum") = 0 Then
								bKeywordPF = True
								bMethodPF = True
								sDetailedError = ""			
							Else
								bKeywordPF = False
								bMethodPF = False
								sDetailedError = "<b>Detailed Description:</b><br>"& Err.Description
							End If					
			Environment("ErrorMsg") = Environment("ErrorMsg") & sDetailedError
	End Sub
