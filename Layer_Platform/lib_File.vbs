
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'objFSO.CopyFile "C:\Pravin_QTP\TestTasks\T14.S529.xls", "C:\Pravin_QTP\TestResultLogs\my1_T14.S529.xls"

'Set fTextFile = objFSO.CreateTextFile("C:\Pravin_QTP\TestTasks\Test.txt",True)
'fTextFile.WriteLine "||| Shree Ganeshai Namaha |||"

Sub subLogDataToFile(sFileName, sLineText)

		Dim objFSO, fLogFile
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set fLogFile = objFSO.OpenTextFile(sFileName,8,True)
		fLogFile.WriteLine(sLineText)

End Sub

Sub subWriteLineToNewFile(sFileName, sLineText)

		Dim objFSO, fLogFile
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set fLogFile = objFSO.OpenTextFile(sFileName,2,True)
		fLogFile.WriteLine(sLineText)

End Sub

Sub subAppendLineToFile(sFileName, sLineText)

		Dim objFSO, fLogFile
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set fLogFile = objFSO.OpenTextFile(sFileName,8,True)
		fLogFile.WriteLine(sLineText)

End Sub