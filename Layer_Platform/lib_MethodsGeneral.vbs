Function getRandomNo(range)

 Dim max,min
 max=split(range,"~")(1)
 min=split(range,"~")(0)
 If max<=min Then
 	getRandomNo=false
 else
	Randomize
 	getRandomNo=Int((max-min+1)*Rnd+min) 
 End If
 
 
End function



Function getRandomLetters(sNumberLength)
	For iCounter = 1 To sNumberLength Step 1
		sval = getRandomNo("65~90")
		sNum = getRandomNo("0~9")
		sString = sString & sNum & chr(sval)
		If len(sString) = sNumberLength Then
			Exit For
		End If
	Next
	getRandomLetters = sString
End Function


Function getRandomAlphaLetters(sNumberLength)
	For iCounter = 1 To sNumberLength Step 1
		sval = getRandomNo("65~90")
		sString = chr(sval)		
	Next
	getRandomLetters = sString
End Function



'**************************************************************************
'    Name: utilSendKeys
'    Purpose: This sub is used to Handle keyboard key strokes

'		Param:  sValue | required
'		AllowedRange: 
'		Description: Value to be sent by keyboard
'
'		Param:  bFlag | required
'		AllowedRange: True/False
'		Description: If True: Will Consider sValue as Special Key. ex:Enter,Home,Ctrl
'					 If False: Will Handle keyboard key strokes normally
'**************************************************************************
Public Sub utilSendKeys(sValue, bFlag)
	
	Dim mySendKeys
 	Set mySendKeys = CreateObject("WScript.shell")
 	
 	If bFlag = True Then
 		sValue = UCase(sValue)
 		sValue = "{"&sValue&"}"
 	End If
 	
 	mySendKeys.SendKeys(sValue)
 	
End Sub



Function GetRepoName(obj)
    GetRepoName = obj.GetTOProperty("TestObjName")
End Function

Function GetClassName(obj)
    GetClassName = obj.GetTOProperty("class Name")
End Function

Function MakeUftName(obj)
    MakeUftName = GetClassName(obj) & "(""" & GetRepoName(obj) & """)"
End Function

Function GetFullUftName(obj)
    dim fullUftName : fullUftName = MakeUftName(obj)
    dim objCurrent : set objCurrent = obj

    do while not IsEmpty(objCurrent.GetTOProperty("parent"))
        set objCurrent = objCurrent.GetTOProperty("parent")
        fullUftName = MakeUftName(objCurrent) & "." & fullUftName
    loop

    GetFullUftName = fullUftName
End Function

'Function is performing regular expression on given string parameter
'Arguments:
'patrn = regExpressionPattern for ex.: "(.*?[^\\])"""
'strng = string to perform regular expression for ex.: "OracleFormWindow("MDM026_OracleFormWindow_MasterItem").OracleTabbedRegion("MDM026_OracleTabbedRegion_Receiving")"
'collection = set True if function should return matches collection, set False when expecting only 1 result
'Example use: RegExpTest("(.*?[^\\])""", "OracleListOfValues(""MDM006_OracleListOfValues_FindBankAccount"")", False)
'Example use: RegExpTest(regExpressionPattern, MakeUftName(TestObject), False)
'Original function can be found in UFT Help,
'TODO: if collection = True - not tested need to be rewritten in order to use
'USE collection = set False to cut characters from repository item name
Function RegExpTest(patrn, strng, collection)
   Dim regEx, Match, Matches, RetStr   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = patrn   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   Set Matches = regEx.Execute(strng)   ' Execute search.
   If collection = True Then
   	   	For Each Match in Matches   ' Iterate Matches collection.
	      RetStr = RetStr & "Match found at position "
	      RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
	      RetStr = RetStr & Match.Value & "'." & vbCRLF
  		Next	
   Else
   		RetStr = Matches.Item(1).Value
   End If
   RegExpTest = RetStr
   Set regEx = Nothing
End Function

'Function reads object from repository and prints them changed in Output tab
' Author: malinraf
'Copy and paste printed objects to repository objects collection as OR_MDM in my case
'Arguments:
'repFolderPath = repository folder and name path from where we want to load the objects for ex.: C:\Projects\AUTOMATION_ATLAS\Layer_Application\OR\MDM.tsr
'testPrefix = set to False when you want to load whole repository, if you want to load particular objects (one test repository) then choose for ex.: MDM006
'Example use: GetRepositoryObjects "C:\Projects\AUTOMATION_ATLAS\Layer_Application\OR\MDM.tsr", False
'Example use: GetRepositoryObjects "C:\Projects\AUTOMATION_ATLAS\Layer_Application\OR\MDM.tsr", "MDM006"
'
Function GetRepositoryObjects(repFolderPath, testPrefix)
	Dim i, shortObjectName, regExpressionPattern
	Set TestObject = Nothing 
	Set ToCollection = Nothing
	Set RepositoryFrom = Nothing
	Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")
	
	RepositoryFrom.Load repFolderPath
	
	Set ToCollection = RepositoryFrom.GetAllObjects
	
	regExpressionPattern = "(.*?[^\\])"""
'	regExpressionPattern = "(.*?[^\\])\(.(.*).{2}" does not work not sure why, it's working in web validator so used Left and Len below to cut right side
	
	If testPrefix = False Then
		For i = 0 To ToCollection.Count-1
			Set TestObject = ToCollection.Item(i)
			shortObjectName = RegExpTest(regExpressionPattern, MakeUftName(TestObject), False)
			shortObjectName = Left(shortObjectName, Len(shortObjectName)-1)
			print "OBJ_"+shortObjectName+" = """+Replace(Cstr(GetFullUftName(TestObject)), """", """""" )+""""
		Next
	Else
		For i = 0 To ToCollection.Count-1
			Set TestObject = ToCollection.Item(i)
			uftObjectName = MakeUftName(TestObject)
			
			If InStr(uftObjectName, testPrefix) <> 0 Then 
				shortObjectName = RegExpTest(regExpressionPattern, uftObjectName, False)
				shortObjectName = Left(shortObjectName, Len(shortObjectName)-1)
				print "OBJ_"+shortObjectName+" = """+Replace(Cstr(GetFullUftName(TestObject)), """", """""" )+""""
			End If	
		Next
	End If
	
	Set TestObject = Nothing 
	Set ToCollection = Nothing 
	Set RepositoryFrom = Nothing
End Function






'   Requires: 
'   sDate (string):   Your Date 
'   sFormat (string):   Your desired format as in "d-m-y, m-d-y, y-m-d" 
'**************************************************************************
Public Function formatDate(sDate, sFormat) 
	sDate = Cstr(sDate)
    sDay    = Split(sDate,"/")(1)
    sMonth    = Split(sDate,"/")(0)
    sYear   = Split(sDate,"/")(2)
        
    Select Case sFormat 
      Case "d-m-y" 
        formatDate = sDay & "/" & sMonth & "/" & sYear 
      Case "m-d-y" 
        formatDate = sMonth & "/" & sDay & "/" & sYear 
      Case "y-m-d" 
        formatDate = sYear & "/" & sMonth & "/" & sDay 
      Case Else 
        formatDate = formatDate
    End Select 
  End Function 
  
 
 
 
 
Function filesDownload(iBrowserCreationTime, sResultPath, sFileName, sFileExtension)
                               On Error Resume Next
                If bKeywordPF = True Then
                                                Browser("CreationTime:=0").SetTOProperty "CreationTime", iBrowserCreationTime
                                               
                                                kyWaitTillTheObjectExist iBrowserCreationTime,OJB_Download_Button_DropDown, 30
                                                kyPerformWebOperation iBrowserCreationTime,OJB_Download_Button_DropDown, "Click", ""                
                                                kyPerformWebOperation iBrowserCreationTime,OBJ_Download_WinMenu,"Select", "Save as"
                                                dtTime = DateAdd("s", 120, now)
                                                While not(kyObjectExistenceCheck (iBrowserCreationTime, OBJ_Download_TextField_FileName, 7)) and dtTime > now 
                                                                kyPerformWebOperation iBrowserCreationTime,OJB_Download_Button_DropDown, "Click", ""
                                                                kyPerformWebOperation iBrowserCreationTime,OBJ_Download_WinMenu,"Select", "Save as"
                                                Wend                                               
'                                               kyWaitTillTheObjectExist iBrowserCreationTime,OBJ_Download_TextField_FileName, 100                                               
                                                kyPerformWebOperation iBrowserCreationTime,OBJ_Download_TextField_FileName, "DblClick", ""

                                                sFullPath = sResultPath & sFileName & "." & sFileExtension
                                                
                                                kyWaitTillTheObjectExist iBrowserCreationTime,OBJ_Download_SaveButton,100
                                                kyPerformWebOperation iBrowserCreationTime,OBJ_Download_TextField_FileName, "Type", sFullPath
                                                kyPerformWebOperation iBrowserCreationTime,OBJ_Download_SaveButton, "Click", ""
                                                
                                                kyWaitTillTheObjectExist iBrowserCreationTime,OBJ_Download_Button_OpenFolder,100                                         
                                                kyPerformWebOperation iBrowserCreationTime, OBJ_Download_Button_Close, "Click", ""
                                                                                                                                                            
                End IF
                  filesDownload = sFullPath 
End Function
  
  
  
  'Syntax: sFolderPath = "D:\AUTOMATION_ATLAS_Rajesh\TestData" ||| Where TestData is the Folder Name that needs to be created  
 Function createFolder(sFolderPath)
  	Set fso=createobject("Scripting.FileSystemObject")
	If fso.FolderExists(sFolderPath) = false Then
	 fso.CreateFolder (sFolderPath)	
	End If	
	Set fso=nothing
 End Function
 
 
 'Syntax: sFolderPath = "D:\AUTOMATION_ATLAS\TestData" ||| Where TestData is the Folder Name that needs to be deleted  
Function deleteFolder(sFolderPath)
	Set fso=createobject("Scripting.FileSystemObject")
	
	If fso.FolderExists(sFolderPath) Then
	 fso.DeleteFolder (sFolderPath)	
	End If		

	Set fso=nothing

End Function




'MALINRAF: function is closing Web Page by it's title(sPageTitle)
Function kyCloseIEPageByTitle (sPageTitle)
	Set oBw=Description.Create
	oBw("micclass").Value="Browser"
	Set oBrws=Desktop.ChildObjects(oBw)
	For i=0 to oBrws.Count-1
		If Browser("micclass:=Browser", "index:=" & i).Exist(0) Then 
		lngBrowserHWND = Browser("micclass:=Browser", "index:=" & i).GetROProperty("hwnd")
		strv= Browser("hwnd:="& lngBrowserHWND).GetROProperty("title")
'		msgbox strv
		Set oBrw=Browser("hwnd:="& lngBrowserHWND)
		Set oPg=Description.Create
		oPg("micclass").Value="Page"
		Set oPage=Browser("hwnd:="& lngBrowserHWND).ChildObjects(oPg)
		For n=0 to oPage.Count-1
			If oPage(n).Exist(0) and Instr(oPage(n).Getroproperty("title"),sPageTitle) <> 0 Then
			oBrw.Close
			Exit For
			End If


		Next
		
		End If
	Next
End Function



Function CostValueConvertToStandardFromNormal(sNormalAmnt)


	asAfterdecimal = Split(sNormalAmnt, ".")
	If UbOund(asAfterdecimal) = 1 Then
		sAfterdecimal = Split(sNormalAmnt, ".")(1)
	Else 
		sAfterdecimal = "00"
	End If
	sBeforeDecimal = asAfterdecimal(0)
	iCount = Len(sBeforeDecimal)
	While iCount <> 0
			iCounter = 1
			While iCounter <= 3 and iCount <> 0
				str = Mid(sBeforeDecimal, iCount, 1) & str
				iCount = iCount - 1
				iCounter = iCounter + 1			
			Wend
			If iCount <> 0 Then
					str = "." & str
			End If

	Wend
	If sAfterdecimal <> "" Then
		CostValueConvertToStandardFromNormal = str & "," & sAfterdecimal
	Else	
		CostValueConvertToStandardFromNormal = str		
	End If
End Function




Function CostValueConvertToNormalFromStandard (sStandardAmnt)
	sAfterDecimal = Split(sStandardAmnt,",")(1)
	sBeforeDecimal = Split(sStandardAmnt,",")(0)
	sStandardAmnt  = Replace(sBeforeDecimal, ".", "")

	If sAfterDecimal <> "" Then
		CostValueConvertToNormalFromStandard = sStandardAmnt & "." & sAfterDecimal
	Else
		CostValueConvertToNormalFromStandard = sStandardAmnt
	End If
End Function


'sDate = dd-mm-yyyy
Function RetrieveMonthFromEnteredDate(sDate, ByRef sYear)
	asDate = Split(sDate, "-")
	sYear = asDate(2)
	sMonth = asDate(1)
	sDay = asDate(0)

		Select Case sMonth
			Case "01"
					RetrieveMonthFromEnteredDate = "Jan"
			Case "02"
					RetrieveMonthFromEnteredDate = "Feb"
			Case "03"
					RetrieveMonthFromEnteredDate = "Mar"
			Case "04"
					RetrieveMonthFromEnteredDate = "Apr"		
			Case "05"
					RetrieveMonthFromEnteredDate = "Mai"
			Case "06"
					RetrieveMonthFromEnteredDate = "Jun"		
			Case "07"
					RetrieveMonthFromEnteredDate = "Jul"
			Case "08"
					RetrieveMonthFromEnteredDate = "Aug"		
			Case "09"
					RetrieveMonthFromEnteredDate = "Sep"
			Case "10" 
					RetrieveMonthFromEnteredDate = "Okt"		
			Case "11"
					RetrieveMonthFromEnteredDate = "Nov"
			Case "12"
					RetrieveMonthFromEnteredDate = "Des"		
		End Select
				If Err.Number <> 0 Then
					bKeywordPF = False
				End If

End Function


'	malinraf
'	------------------------------------------------------------------------
'	numbersFromStringExtract
'		Description: 
'		This function extracts number chars from defined string.
'
'		Parameter 'firstNumericChainOnly' indicates whether number
'		extraction should stop when first non-number char occurs after
'		previous numeric char (this prevents from unwanted chaining 
'		of multiple numbers)
'		Parameter 'startText' indicates on which character number
'		extraction should start
'		Parameter 'stopText' indicates on which character number
'		extraction should Stop
'	------------------------------------------------------------------------
Function numbersFromStringExtract(defString, firstNumericChainOnly, startText, stopText, sDelimiter)
	extraction = ""
	
	If startText <> "" Then
		startPos = InStr(1, defString, startText)
		defString = Mid(defString, startPos, Len(defString))
	End If
	
	If stopText <> "" Then
		stopPos = InStr(1, defString, stopText)
	Else
		stopPos = ""
	End If
	
	If firstNumericChainOnly Then
'	first numeric chain only
		numberFound = false
		For i = 1 To Len(defString)
			curChar = Mid(defString,i,1)
			If isNumeric(curChar) Then
				numberFound = true
				extraction = extraction & curChar
			Else
				If numberFound Then
					Exit For
				End If
			End If		
		Next
	
	Else
	
'	all numeric chars
		If stopText <> "" Then
			sIterator = stopPos
		Else
			sIterator = Len(defString)
		End If
		For i = 1 To sIterator

			curChar = Mid(defString,i,1)
			If isNumeric(curChar) Then
				extraction = extraction & curChar
			ElseIf curChar = sDelimiter Then
				extraction = extraction & curChar
			End If		
		Next
	
	End If	
	numbersFromStringExtract = extraction
	
End Function


Function StringSetMapAsDictionary(sFirstString, sSecondString, sDelimiter)
Dim asFirstString, asSecondString, iCounter

	Set oDataSetDictionary = CreateObject("Scripting.Dictionary")
	
	'Split the arrays with the delimiter
	asFirstString = Split(sFirstString, sDelimiter)
	asSecondString = Split(sSecondString, sDelimiter)
	
	
			If Ubound(asFirstString) = Ubound(asSecondString) Then
				For iCounter = Lbound(asFirstString) To Ubound(asFirstString) Step 1
					oDataSetDictionary.Add asFirstString(iCounter), asSecondString(iCounter)
				Next
					Set	StringSetMapAsDictionary = oDataSetDictionary
			Else
					Set	StringSetMapAsDictionary = ""
			End If
	
End Function
