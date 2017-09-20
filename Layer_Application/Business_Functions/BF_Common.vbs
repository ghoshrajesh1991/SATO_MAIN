'***********************************************************************************************
'Name: sub_SalesOrderNumberFetch
'Creator: Rajesh Ghosh
'Purpose: This method Fetches the Sales order Number from the Status bar
'***********************************************************************************************
Sub sub_ToolBarMenuNavigate(sNavigation)
	on Error Resume Next
		 If bMethodPF = True Then
			kyPerformSAPOperation OBJ_Comm_SAPGuiMenuBar_Window_menubar, "Select", sNavigation
			kyPerformSAPOperation OBJ_Comm_SAPGuiStatusBar_Window_StatusBar, "Sync", ""
								passed = "Navigate to: <b>" &sNavigation & "</b>"								
								failed= "Error while Navigating to: <b>" &sNavigation & "</b>"	
								subWriteTestStepResult passed, failed, ""
		 End If
End Sub


Sub sub_LaunchSAPApplication()

	sapCloseAllInstance
	On Error Resume Next	
			If bMethodPF Then
						InvokeApplication Universal.Item("SAPLocation")
						kySAPWaitTillObjectExist OBJ_PO001_WinListView_SAPStartScreen_Servers, 50
						kyPerformSAPOperation OBJ_PO001_Dialog_SAPStartScreen, "Maximize", ""
								passed = "Launch SAP ERP Application"								
								failed= "Error while launching SAP ERP."
								subWriteTestStepResult passed, failed, ""
		
			End If	

End Sub


Sub sub_UserLogin()
	on Error Resume Next
	 If bMethodPF = True Then
	
	 	 kyPerformSAPOperation OBJ_PO001_SAPGuiEdit_SAP_Client, "Set", UNIVERSAL.Item("Client")
		 kyPerformSAPOperation OBJ_PO001_SAPGuiEdit_SAP_User, "Set", UNIVERSAL.Item("UserName")
		 kyPerformSAPOperation OBJ_PO001_SAPGuiEdit_SAP_Password, "Set", UNIVERSAL.Item("Password")	 
		 kyPerformSAPOperation OBJ_PO001_SAPGuiEdit_SAP_LogonLanguage, "Set", UNIVERSAL.Item("LogonLanguage")
		 
				passed="Set the Logon parameters<br>"&_
					   "Client:<b>"&UNIVERSAL.Item("Client")&"</b><br>"&_
					   "UserName:<b>"&UNIVERSAL.Item("UserName")&"</b><br>"&_
					   "LogonLanguage:<b>"&UNIVERSAL.Item("LogonLanguage")&"</b>"
				failed="Error while setting the logon parameters"
				subWriteTestStepResult passed,failed, true
				
		kyPerformSAPOperation OBJ_PO001_SAPGuiButton_SAP_Enter, "Click", ""
		kySAPWaitTillObjectExist OBJ_PO001_SAPGuiTree_SAPEasyAccess_TableTreeControl, 50

				passed="Login successfull"
				failed="Error while Login to the sever<br>"&_
						"Please check the data in the Login page"
				subWriteTestStepResult passed, failed, true
		
	 End If 
End Sub


Sub sub_SelectServer()
	on Error Resume Next
	 If bMethodPF = True Then
		 kyPerformSAPOperation OBJ_PO001_WinListView_SAPStartScreen_Servers, "Select", UNIVERSAL.Item("ServerName")
		 kyPerformSAPOperation OBJ_PO001_WinButton_SAPStartScreen_LogOn, "Click", ""
		 kySAPWaitTillObjectExist OBJ_PO001_SAPGuiEdit_SAP_User, 50
		 kyPerformSAPOperation OBJ_PO001_Dialog_SAPStartScreen, "Minimize", ""
		 kyPerformSAPOperation OBJ_PO001_SAPGuiWindow_SAP, "Activate", ""
		 kyPerformSAPOperation OBJ_PO001_SAPGuiEdit_SAP_User, "Highlight", ""

				passed="Select Server <b><br>"&Universal.Item("ServerName")&"</b>"
				failed="Error while selecting the server <b> "&Universal.Item("ServerName")&"</b> "
				subWriteTestStepResult passed,failed, true	
		
	 End If 
End Sub

Sub sapCloseAllInstance()
				On Error Resume Next	
					If bMethodPF Then
						subInitializeFunctionVariables()
						kySAPInstanceClose()
					End If
					
									
End Sub

Sub sub_CloseSAPERP()
				On Error Resume Next	

						kySAPInstanceClose()
						
						passed = "Close SAP ERP application"
						failed="Error while closing SAP ERP application"
						subWriteTestStepResult passed,failed, ""
			
End Sub


Sub sub_PageNavigate(sTransactionCode, sExpPageObject, sPageDesc)
	on Error Resume Next
	 If bMethodPF = True Then
	 
	 dtTime = DateAdd("s",100,now)
	 Do
	 	bFLag = kySAPIfObjectExistsCheck (OBJ_PO001_SAPGuiOKCode_SAPEasyAccess_OKCode, 4)
			If bFlag = False Then
				kyPerformSAPOperation OBJ_PO001_SAPGuiButton_CreateArticleHierarchy_Cancel, "Click", ""
				kySAPWaitTillObjectExist OBJ_PO001_SAPGuiOKCode_SAPEasyAccess_OKCode, 10
			Else
				Exit Do
			End If
	 Loop While dtTime > now and bKeywordPF = True
	 	
	 
		 kyPerformSAPOperation OBJ_PO001_SAPGuiOKCode_SAPEasyAccess_OKCode, "Set", sTransactionCode
		 kyPerformSAPOperation OBJ_PO001_SAPGuiButton_SAPEasyAccess_Enter, "Click", ""
		 kySAPWaitTillObjectExist sExpPageObject, 50
		 kyPerformSAPOperation sExpPageObject, "Highlight", ""
		 
	 			passed = "Navigate to <b>" & sPageDesc
				failed="Error while navigating to <b>" & sPageDesc
				subWriteTestStepResult passed,failed, true

		
	 End If 
End Sub


Sub sub_SAPLogOff()
	on Error Resume Next
	 If bMethodPF = True Then
	 
	 		kyPerformSAPOperation OBJ_PO001_SAPGuiMenubar_CreateArticleHierarchy_menubar, "Select", "System;Log Off"
	 			passed = "Logoff from the Server"
				failed="Error while logging off from the server"
				subWriteTestStepResult passed,failed, true
	 		kyPerformSAPOperation OBJ_PO001_SAPGuiButton_LogOff_Yes, "Click", ""
	 		

	 End If

End Sub
