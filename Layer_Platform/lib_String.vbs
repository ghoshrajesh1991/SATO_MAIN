Function kySearchString(textToBeSearch, str)
	
		On Error Resume Next		
		 If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim stringFoundCode
	
					stringFoundCode=instr(textToBeSearch,str)
				
					If Err.Number = 0 AND Err.Description = "" AND stringFoundCode <> 0 Then
							bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kySearchString=stringFoundCode
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyFormatAmount'  " & Err.Description	
							kySearchString= stringFoundCode
					End If
			
		'**********************************************************************************************************	
End if
					
			
	End Function
