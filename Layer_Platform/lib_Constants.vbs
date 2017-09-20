
' =========  Veriables  =========
Dim PassBitmapCapture: PassBitmapCapture = True
Dim FailBitmapCapture: FailBitmapCapture = True
Dim bKeywordPF
Dim dummyUserID
Dim dummyUserPassword
Dim bMethodPF
Dim sKeywordError
Dim sMethodError
Dim bTestCasePF
Dim sTestCaseError
Dim sTestCaseMessage
Dim sDelimeter
Dim iBrowserCreationTimeGlobal
Dim iFunctionIteration
Dim sTestFlow
Dim bResultIndicator
Dim OBJ
Dim sSheetName
Dim sModuleSheetName
Dim iIterationRow
Dim sFullExcelPath
Dim sTestCaseName
Dim sDataSetColumn
Set OBJ = CreateObject("Scripting.Dictionary")
Set oExcelDictionary = CreateObject("Scripting.Dictionary")
Dim EXEC
Dim UNIVERSAL
        Set UNIVERSAL = CreateObject ("Scripting.Dictionary")
        Set EXEC = CreateObject("Scripting.Dictionary")	
Dim oDataSetDictionary
Dim oDicCommonDataRepo
		Set oDicCommonDataRepo = CreateObject ("Scripting.Dictionary")

	
	
	
'==========project setup
Dim ORType
ORType="ModuleWise"  ' allowed values:   "ProjectLevel", "ModuleWise"
Dim mController
mController="SAP_Modules.xls"

Dim projectName
projectName="SAP AUTOMATION"

Dim projectTagLine
projectTagLine="TAG LINE FOR THE EBS OR ANY OTHER STATMENT"

Dim projectOR
projectOR= "project.tsr"

Dim ORFolderPath
'ORFolderPath="C:\AUTOMATION_ATLAS\Layer_Application\OR"
ORFolderPath=Split(Environment.Value("TestDir"),"Layer_Test")(0) & "Layer_Application\OR"
Dim CurrentExcelSheetPath



Dim takeSnaps
takeSnaps="All"   'allowed values:  "All" "OnFail"

Dim synTime
synTime=20  'seconds

Dim OpenResultAfterRun
OpenResultAfterRun="No"   'allowed values "Yes"  "No"

Dim eStart
Dim mStart
Dim tStart
Dim eEnd
Dim mEnd
Dim tEnd


' =========  Veriables  Default Initialization=========

iBrowserCreationTimeGlobal = 0
iFunctionIteration = 1
sTestCaseMessage = ""

' =========  Veriables  Default Initialization=========

'Dim  sTestTasksFolderPath
'Dim  sTestDataFolderPath
'Dim  sTestResultLogFolderPath
Dim cDR, cDRTABLE
Dim cTESTREPORTLOGO
Dim  cTESTCONFIGFOLDERPATH
Dim  cTESTDATAFOLDERPATH
Dim  cTESTRESULTLOGFOLDERPATH
Dim cLISEXESHEET
Dim cTABLEMODULES
Dim cROOTPATH
Dim iPassCount
Dim iFailCount
Dim iStepCount
iStepCount=0

cROOTPATH = Split(Environment.Value("TestDir"),"Layer_Test")(0)

'''''''''''sTestConfigFolderPath =  			Split(Environment.Value("TestDir"),"TestScripts")(0) & "TestConfig\"
'''''''''''sTestDataFolderPath = 			 	  Split(Environment.Value("TestDir"),"TestLayer")(0) & "TestData\"
'''''''''''sTestResultLogFolderPath =   	Split(Environment.Value("TestDir"),"TestLayer")(0) & "TestResultLogs\"
cTESTREPORTLOGO		=					Split(Environment.Value("TestDir"),"TestScripts")(0) & "Logo"
cTESTCONFIGFOLDERPATH = 				Split(Environment.Value("TestDir"),"TestScripts")(0) & "TestConfig\"
cTESTDATAFOLDERPATH   = 			 	  	Split(Environment.Value("TestDir"),"Layer_Test")(0) & "TestData\"
cTESTRESULTLOGFOLDERPATH =   		Split(Environment.Value("TestDir"),"Layer_Test")(0) & "TestResultLogs\"
cLISEXESHEET = "LIS_Execution.xls"
cTABLEMODULES = "tbl_modules"
			cDR = "DataRepository.xls"
			cDRTABLE = "Table1"
sDelimeter = "::"

'========System variable

Dim systemDate,systemDate_mmSddSyyyy

' =========  wait variable  =========
Dim w1,w2,w3,w4,w5,w6,w7,w8,w9,w10,w11,	w12,w13,w14,w15,w16,w17,w18,w19,w20,w21,w22,w23,w24,w25,w26,w27,w28,w29,w30,w31,w32,w33,w34,w35,w36,w37,w38,w39,w40,w41,w42,w43,w44,w45,w46,w47,w48,w49,w50,w51,w52,w53,w54,w55,w56,w57,w58,w59,w60

w1=1
w2=2
w3=3
w4=4
w5=5
w6=6
w7=7
w8=8
w9=9
w10=10
w11=11
w12=12
w13=13
w14=14
w15=15
w16=16
w17=17
w18=18
w19=19
w20=20
w21=21
w22=22
w23=23
w24=24
w25=25
w26=26
w27=27
w28=28
w29=29
w30=30
w31=31
w32=32
w33=33
w34=34
w35=35
w36=36
w37=37
w38=38
w39=39
w40=40
w41=41
w42=42
w43=43
w44=44
w45=45
w46=46
w47=47
w48=48
w49=49
w50=50
w51=51
w52=52
w53=53
w54=54
w55=55
w56=56
w57=57
w58=58
w59=59
w60=60




