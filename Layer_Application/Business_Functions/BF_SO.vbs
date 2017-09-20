'***********************************************************************************************
'Name: sub_SalesOrderStatusVerify
'Creator: Rajesh Ghosh
'Param: sTestObjectName|sExpectedText  The Field object that needs to be Checked
		'Text or Status that is to be verified
'Purpose: This method verifies the Status 
'***********************************************************************************************
Sub sub_SalesOrderStatusVerify(sTestObjectName, sFieldName, sExpectedText)
	on Error Resume Next
		 If bMethodPF = True Then	
		 
		 	kySAPCompareObjectProperty sTestObjectName, "text", sExpectedText
	 			passed= "Verify the '" & sFieldName & "' is set to as <b>" & sExpectedText
				failed= "Error Verifying the Status of '" & sFieldName & "' is set to as <b>" & sExpectedText
				subWriteTestStepResult passed, failed, true
				
	 	 End If
End Sub




'***********************************************************************************************
'Name: sub_SalesOrderNumberDisplay
'Creator: Rajesh Ghosh
'Param: iOrderNumber  The order number that needs to be displayed
'Purpose: This method Opens the Order Number
'***********************************************************************************************
Sub sub_SalesOrderNumberDisplay(iOrderNumber)
	on Error Resume Next
		 If bMethodPF = True Then

			kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_DisplaySalesDoc_Order, "Set", iOrderNumber
	 		kyPerformSAPOperation OBJ_Comm_SAPGuiButton_Window_Continue, "Click", ""
	 		kyStatusBarSync
	 		
	 			passed= "Open Order Number: <b>"& iOrderNumber
				failed= "Error while opening Order Number: <b>"& iOrderNumber
				subWriteTestStepResult passed, failed, true
	 	 End If
End Sub



'***********************************************************************************************
'Name: sub_SalesOrderNumberFetch
'Creator: Rajesh Ghosh
'Purpose: This method Fetches the Sales order Number from the Status bar
'***********************************************************************************************
Sub sub_SalesOrderNumberFetch()
	on Error Resume Next
		 If bMethodPF = True Then
	 		
	 		'Click on the Information Pop-up if exist
	 		bFLag = kySAPIfObjectExistsCheck (OBJ_Comm_SAPGuiWindow_Information_Continue, 10)
			If bFlag Then
				kyPerformSAPOperation OBJ_Comm_SAPGuiWindow_Information_Continue, "Highlight", ""
				kyPerformSAPOperation OBJ_Comm_SAPGuiWindow_Information_Continue, "Click", ""
			End If
			kyStatusBarSync
			kySAPSync OBJ_Comm_SAPGuiStatusBar_Window_StatusBar, 4
	 		
	 		'Fetch the Sales Order Number generated in the Status Bar
	 		 sStatusBarMsg = kyGetSAPObjectProperty (OBJ_Comm_SAPGuiStatusBar_Window_StatusBar, "text")
	 		 UNIVERSAL.Item("SO_Number") = numbersFromStringExtract(sStatusBarMsg, True, "POS", "", "")
	 		
	 			passed= "Fetch the Sales Order number"
				failed= "Error while fethching Sales Order Number"
				subWriteTestStepResult passed, failed, true
				
				kyExcelDataExport CurrentExcelSheetPath, sSheetName, iIterationRow, "Sales_Order_No", UNIVERSAL.Item("SO_Number")
	 	 End If
End Sub



'***********************************************************************************************
'Name: sub_createSalesOrder
'Creator: Rajesh Ghosh
'Purpose: This method intend to create Sales Order
'***********************************************************************************************
Sub sub_createSalesOrder()
	on Error Resume Next
	 If bMethodPF = True Then
	 	kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_CreateSales_OrderType, "Set", UNIVERSAL.Item("Order Type")
	 	kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_CreateSales_SalesOrg, "Set", UNIVERSAL.Item("Sales Organization")
	 	kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_CreateSales_DistrCh, "Set", UNIVERSAL.Item("Distribution Channel")
	 	kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_CreateSales_Division, "Set", UNIVERSAL.Item("Division")
	 		 			passed= "Fill Organization data."
						failed="Error while filling Organization data"
						subWriteTestStepResult passed, failed, true
	 	kyPerformSAPOperation OBJ_Comm_SAPGuiButton_Window_Continue, "Click", ""
	 	
	 	sub_CustomerDetailsInHeaderFill
	 	
	 	For iCounter = 1 To UNIVERSAL.Item("No of Items") Step 1
	 		sub_FillItemDetails iCounter
	 	Next
 			
 			passed= "Fill Sales Order Document"
			failed= "Error while filling Sales Order Document"
			subWriteTestStepResult passed, failed, true
			
	 	kyPerformSAPOperation OBJ_Comm_SAPGuiButton_Window_Save, "Click", ""
	 		passed= "Save the Sales Order"
			failed= "Error while saving sales order"
			subWriteTestStepResult passed, failed, true
	 End If
	
End Sub	
	 
'***********************************************************************************************
'Name: sub_SalesOrderConfirmBeforSave
'Creator: Rajesh Ghosh
'Purpose: This method Clicks on Complete Dlv button to confirm the Sales order
'***********************************************************************************************
Sub sub_SalesOrderConfirmBeforSave()
	on Error Resume Next
		 If bMethodPF = True Then
	 
			 kyPerformSAPOperation OBJ_SO001_SAPGuiButton_CreateSales_Completedlv, "Click", ""
	 			passed= "Click on Complete Dlv to confirm"
				failed= "Error while Clicking on Complete Dlv"
				subWriteTestStepResult passed, failed, true
	 	 End If
End Sub
	 	
'***********************************************************************************************
'Name: sub_CustomerDetailsInHeaderFill
'Creator: Rajesh Ghosh
'Purpose: This method intend to fill the Customer details while creation of Sales Order
'***********************************************************************************************
Sub sub_CustomerDetailsInHeaderFill
	on Error Resume Next
		 If bMethodPF = True Then
	
			kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_CreateSales_SoldToParty, "Set", UNIVERSAL.Item("SoldToParty")
		 	kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_CreateSales_ShipToParty, "Set", UNIVERSAL.Item("ShipToParty")
		 	kyPerformSAPOperation OBJ_SO001_SAPGuiEdit_CreateSales_CustReference, "Set", UNIVERSAL.Item("Customer Reference")
		 				passed= "Fill cutomer details"
						failed="Error while filling customer details."
						details = "Customer Name: "&UNIVERSAL.Item("SoldToParty")&_
								  "<br>Cutomer Ref: "& UNIVERSAL.Item("Customer Reference")
						subWriteTestStepResult passed, failed, details
						
		End If
End Sub
	
	 
'***********************************************************************************************
'Name: sub_FillItemDetails
'Creator: Rajesh Ghosh
'Param: Row number
'Purpose: This method intend to fill the Item details while creation of Sales Order
'Syntax: sub_FillItemDetails 1
'***********************************************************************************************
 	Sub sub_FillItemDetails(iRowID)
	 	on Error Resume Next
		 If bMethodPF = True Then
'		 			kyPerformSAPOperation OBJ_SO001_SAPGuiTable_CreateSales_Allitems, "SelectCell ", "4,Description"
			 		kySAPTableFill OBJ_SO001_SAPGuiTable_CreateSales_Allitems, iRowID, UNIVERSAL.Item("Item Fields") , UNIVERSAL.Item("Item Values" & iRowID)
						passed= "Materials Details filled in <b>" & iRowID & "</b> row" 
						failed="Error while filling Materials details.<br>"&_
								"Please re-verify the Item table."
						subWriteTestStepResult passed,failed, ""	
						
		End If
 	End Sub
	 		

