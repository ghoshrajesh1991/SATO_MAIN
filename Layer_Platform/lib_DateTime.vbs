
Function mGetTimeStampLong1()
	 mGetTimeStampLong1 = Replace(Replace(FormatDateTime (Date,1),",","")," ","_")  & "_" & Replace( FormatDateTime(Time,3),":","_")
End Function


Function mGetTimeStampLong()
	'mGetTimeStampLong = FormatDateTime (Date,1) & " / " & FormatDateTime(Time,3)
	mGetTimeStampLong=date&"_"& Hour(Now)&"."&Minute(Now)&"."&Second(Now)

End Function

	Function mGetDate(sInterval, iNumber, sEnv)

			Dim sDate
			sDate = DateAdd(sInterval, iNumber, date)
			mGetDate = CStr(mFormatEnvironmentDate(sEnv, sDate) )'DatePart("d",sDate) & "." & DatePart("m",sDate) & "." & DatePart("yyyy",sDate)))

	End Function

	Function mFormatEnvironmentDate(sEnvironment, sDate)

			Select Case UCase(sEnvironment)
			
					Case "FI"
									mFormatEnvironmentDate = CStr(mGetTwoDigitChar(DatePart("d",CDate(sDate))) & "." & mGetTwoDigitChar(DatePart("m",CDate(sDate))) & "." & DatePart("yyyy",CDate(sDate)))

					Case "SE"
									mFormatEnvironmentDate = CStr(DatePart("yyyy",CDate(sDate)) & "-" & mGetTwoDigitChar(DatePart("m",CDate(sDate))) & "-" & mGetTwoDigitChar(DatePart("d",CDate(sDate))) )			

					Case Else
									mFormatEnvironmentDate = CStr(mGetTwoDigitChar(DatePart("d",CDate(sDate))) & "." & mGetTwoDigitChar(DatePart("m",CDate(sDate))) & "." & DatePart("yyyy",CDate(sDate)))
			
			End Select

	End Function


	Function mGetTwoDigitChar(sStr)

			If CInt(sStr) < 10 Then
						mGetTwoDigitChar = "0" & sStr
			Else
						mGetTwoDigitChar =  sStr
			End If
			
	End Function
	
	Function kyGetVBSFormattedDate(dateToFormat)
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim formattedDate
			If (IsDate(dateToFormat)) Then			
					d=Day(dateToFormat)
					m=Month(dateToFormat)
					y=year(dateToFormat)
					
					Dim dateToFormat1
					dateToFormat1 = CDate(mGetTwoDigitChar(d) & "-"& mGetTwoDigitChar(m) &"-"& y )  ' As system gives date in ddmmyyyy format but VBS consdier date as mmddyyyy format hence swaping done before formatting
					
					d=Day(dateToFormat1)
					m=Month(dateToFormat1)
					y=year(dateToFormat1)
					
					formattedDate = CStr(mGetTwoDigitChar(m) & "/"& mGetTwoDigitChar(d) &"/"& y )
					
		
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyGetVBSFormattedDate=formattedDate
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyGetVBSFormattedDate'  " & Err.Description	
							kyGetVBSFormattedDate= sKeywordError
					End If
			Else
							sKeywordError = "Error in 'kyGetVBSFormattedDate'"
			End If
		'**********************************************************************************************************	
		End if 
					
			
	End Function
	
	'momin
	Function kyGetFormattedDate(dateToFormat, formatOfDate) ' dateToFormat should always be in VBS supported format VBS support first month then day and then year
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim formattedDate
			If (IsDate(dateToFormat)) Then			
					d=Day(dateToFormat)
					m=Month(dateToFormat)
					y=year(dateToFormat)
					
						
'					dateToFormat = CStr(mGetTwoDigitChar(d) & "-"& mGetTwoDigitChar(m) &"-"& y )  ' As system gives date in ddmmyyyy format but VBS consdier date as mmddyyyy format hence swaping done before formatting
'					
'					d=Day(dateToFormat)
'					m=Month(dateToFormat)
'					y=year(dateToFormat)
					
					Select Case formatOfDate
		
						Case "ddmmyy"
							formattedDate = CStr(mGetTwoDigitChar(d) & mGetTwoDigitChar(m) & right(y,2) )							
						Case "dd-mm-yyyy"
							formattedDate = CStr(mGetTwoDigitChar(d)&"-" & mGetTwoDigitChar(m)&"-" & y )
						Case "yyyymmdd"
							formattedDate = CStr(y & mGetTwoDigitChar(m) &mGetTwoDigitChar(d))
						Case "d-mm-yy"
							formattedDate = CStr(d & "-"& mGetTwoDigitChar(m) &"-"& right(y,2) )
						Case "mm/dd/yyyy"
							formattedDate = CStr(mGetTwoDigitChar(m) & "/"& mGetTwoDigitChar(d) &"/"& y )
						Case "dd/mm/yyyy"
							formattedDate = CStr(mGetTwoDigitChar(d) & "/"& mGetTwoDigitChar(m) &"/"& y )
						
						Case Else
							
					End Select			
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyGetFormattedDate=formattedDate
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyGetFormattedDate'  " & Err.Description	
							kyGetFormattedDate= sKeywordError
					End If
			Else
							sKeywordError = "Error in 'kyGetFormattedDate'"
			End If
		'**********************************************************************************************************	
		End if 
					
			
	End Function
	
		'momin
	Function kyAddDaysToTheDate(srcDate, noOfDaysTobeAdded)
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim newdate
			
			If (IsDate(srcDate)) Then			
					
					newdate=DateAdd("d", cint (noOfDaysTobeAdded), srcDate)					
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyAddDaysToTheDate=newdate
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyAddDaysToTheDate'  " & Err.Description	
							kyAddDaysToTheDate= sKeywordError
					End If
			Else
							sKeywordError = "Error in 'kyAddDaysToTheDate'"
			End If
		'**********************************************************************************************************	
		End if 
		
		
	End Function
	
		Function kyAddYearsToTheDate(srcDate, noOfYearsTobeAdded)
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim newdate
			
			If (IsDate(srcDate)) Then			
					
					newdate=DateAdd("yyyy", cint (noOfYearsTobeAdded), srcDate)					
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyAddYearsToTheDate=newdate
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyAddYearsToTheDate'  " & Err.Description	
							kyAddYearsToTheDate= sKeywordError
					End If
			Else
							sKeywordError = "Error in 'kyAddYearsToTheDate'"
			End If
		'**********************************************************************************************************	
		End if 
			
	End Function
	
	
	
	Function kyGetDayFromDate(srcDate)
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim dayFromDate
			
			If (IsDate(srcDate)) Then			
					
					dayFromDate=Day(srcDate)					
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyGetDayFromDate=dayFromDate
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyGetDayFromDate'  " & Err.Description	
							kyGetDayFromDate= sKeywordError
					End If
			Else
							sKeywordError = "Error in 'kyGetDayFromDate'"
			End If
		'**********************************************************************************************************	
		End if 
			
	End Function
	
	
	Function kyGetMonthFromDate(srcDate)
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim monthFromDate
			
			If (IsDate(srcDate)) Then			
					
					monthFromDate=Month(srcDate)					
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyGetMonthFromDate=monthFromDate
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyGetMonthFromDate'  " & Err.Description	
							kyGetMonthFromDate= sKeywordError
					End If
			Else
							sKeywordError = "Error in 'kyGetMonthFromDate'"
			End If
		'**********************************************************************************************************	
		End if 
			
	End Function
	
	
	
		Function kyGetYearFromDate(srcDate)
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim yearFromDate
			
			If (IsDate(srcDate)) Then			
					
					yearFromDate=Year(srcDate)					
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyGetYearFromDate=yearFromDate
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyGetYearFromDate'  " & Err.Description	
							kyGetYearFromDate= sKeywordError
					End If
			Else
							sKeywordError = "Error in 'kyGetYearFromDate'"
			End If
		'**********************************************************************************************************	
		End if 
			
	End Function
	
	
	'momin
	Function kyGetTimeStamp() ' dateToFormat should always be in VBS supported format VBS support first month then day and then year
	
		On Error Resume Next		
		If bKeywordPF = True Then' If hte Previous Keyword is Passed then Execute the Code - Else No Action
		'*************************************************************************************************

			Dim timeStamp,tempTimeStamp
			tempTimeStamp=now()
						
					d=Day(tempTimeStamp)
					m=Month(tempTimeStamp)
					y=year(tempTimeStamp)
					hr=hour(tempTimeStamp)
					min=minute(tempTimeStamp)
					se=second(tempTimeStamp)
					
					timeStamp= CStr(y & mGetTwoDigitChar(m) & mGetTwoDigitChar(d) & mGetTwoDigitChar(hr) & mGetTwoDigitChar(min) & mGetTwoDigitChar(se))
				
					If Err.Number = 0 AND Err.Description = ""  Then
							'bKeywordPF = true' Keyword Objective is Achived so value is Set to True
							kyGetTimeStamp=timeStamp
					Else
							bKeywordPF = false
							sKeywordError = "Error in 'kyGetTimeStamp'  " & Err.Description	
							kyGetTimeStamp= sKeywordError
					End If
			
		'**********************************************************************************************************	
		End if 
					
			
	End Function
	
	
	
Function SplitSec(pNumSec)
  Dim d, h, m, s
  Dim h1, m1

  d = int(pNumSec/86400)
  h1 = pNumSec - (d * 86400)
  h = int(h1/3600)
  m1 = h1 - (h * 3600)
  m = int(m1/60)
  s = m1 - (m * 60)
  SplitSec = cStr(d) & "D:" & cStr(h) & "H:" & cStr(m) & "M:" & cStr(s)& "S" 
  
End Function
