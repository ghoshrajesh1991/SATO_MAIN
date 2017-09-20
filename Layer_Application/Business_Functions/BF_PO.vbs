




Sub Sub_ModifyArticleNode()
	on Error Resume Next
	 If bMethodPF = True Then
	 
	 	kyPerformSAPOperation OBJ_PO001_SAPGuiEdit_CreateArticleHierarchy_ChangeHierarchy, "Set", UNIVERSAL.Item("HierarchyName")
	 	kyPerformSAPOperation OBJ_PO001_SAPGuiButton_CreateArticleHierarchy_Execute, "Click", ""
	 	
	 	
	 	kyPerformSAPOperation OBJ_PO001_SAPGuiToolbar_CreateArticleHierarchy_GridToolbar, "PressButton", "NODECREATE"
		kyFillTableRowByRowAndColID OBJ_PO001_SAPGuiGrid_CreateArticleHierarchy_GridViewCtrl, 1, "Hierarchy Node",UNIVERSAL.Item("SubChildNode")
		kyFillTableRowByRowAndColID OBJ_PO001_SAPGuiGrid_CreateArticleHierarchy_GridViewCtrl, 1, "Description",UNIVERSAL.Item("SubChildNodeDesc")
		
				passed = "Add Child Node to the Parent: <b>"&UNIVERSAL.Item("ChildNode")& "</b> node"
				failed= "Error while adding child node: <b>"&UNIVERSAL.Item("SubChildNode")&"</b>"
				subWriteTestStepResult passed,failed, true
		
		kyPerformSAPOperation OBJ_PO001_SAPGuiTabStrip_CreateArticleHierarchy_MAINTENANCE, "Select", "ArticleAssignments"
		kyPerformSAPOperation OBJ_PO001_SAPGuiToolbar_CreateArticleHierarchy_GridToolbar, "PressButton", "CELLINSERT"
		kyFillTableRowByRowAndColID OBJ_PO001_SAPGuiGrid_CreateArticleHierarchy_GridViewCtrl, 1,"Material", UNIVERSAL.Item("Material")
		

				passed= "Material: <b>"&UNIVERSAL.Item("Material")&"</b> linked successfully to Parent node: <b>"&UNIVERSAL.Item("ChildNode")& "</b>"
				failed="Error while linking Material:  <b>"&UNIVERSAL.Item("Material")&"</b> to Parent node: <b>"&UNIVERSAL.Item("ChildNode")& "</b>"
				subWriteTestStepResult passed,failed, true
		
		kyPerformSAPOperation OBJ_PO001_SAPGuiButton_CreateArticleHierarchy_Save, "Click", ""
		kySAPSync OBJ_PO001_SAPGuiStatus_CreateArticleHierarchy_StatusBar, 5
	 	kySAPCompareObjectProperty OBJ_PO001_SAPGuiStatus_CreateArticleHierarchy_StatusBar, "text", "Article hierarchy "& UNIVERSAL.Item("HierarchyName") &" was saved"
	 			passed = "Verify the Hierarchy is modified"
				failed="Error while modifying and saving hierarchy :<b>"&UNIVERSAL.Item("HierarchyName")&"</b>"
				subWriteTestStepResult passed,failed, true
	 End If
	 				'********************************************************************************************
				'subMethodVerification EXEC, "Sub_ModifyArticleNode", "Method Message"
End Sub


Sub Sub_HierachySaveVerify()
		on Error Resume Next
	 If bMethodPF = True Then
	 	kySAPCompareObjectProperty OBJ_PO001_SAPGuiStatus_CreateArticleHierarchy_StatusBar, "text", "Article hierarchy "& UNIVERSAL.Item("HierarchyName") &" was saved"
	 			passed = "Verify the New Hierarchy is saved"
				failed="Error while saving hierarchy :<b>"&UNIVERSAL.Item("HierarchyName")&"</b>"
				subWriteTestStepResult passed,failed, true
	 End If
	 				'********************************************************************************************
				'subMethodVerification EXEC, "Sub_HierachySaveVerify", "Method Message"
End Sub


Sub Sub_CreateArticleNode()
	on Error Resume Next
	 If bMethodPF = True Then
	 	
	 	 sub_PageNavigate "/nwmatgrp01", OBJ_PO001_SAPGuiEdit_CreateArticleHierarchy_Hierarchy, "Create Article Hierarchy"
	 	 
	 	 kyPerformSAPOperation OBJ_PO001_SAPGuiEdit_CreateArticleHierarchy_Hierarchy, "Set", UNIVERSAL.Item("HierarchyName")
	 	 kySAPCheckBoxToggle OBJ_PO001_SAPGuiEdit_CreateArticleHierarchy_SAPBW, "UNCHECK"
	 	 kyFillTableRowByRowAndColID OBJ_PO001_SAPGuiTable_CreateArticleHierarchy_LangMain, 1, "Lang.","EN"
	 	 kyFillTableRowByRowAndColID OBJ_PO001_SAPGuiTable_CreateArticleHierarchy_LangMain, 1, "Description",UNIVERSAL.Item("HierarchyDesc")

				passed = "Enter article hierarchy details "
				failed="Error while filling hierarchy details"
				subWriteTestStepResult passed,failed, true
				
		kyPerformSAPOperation OBJ_PO001_SAPGuiButton_CreateArticleHierarchy_Execute, "Click", ""
		bFLag = kySAPIfObjectExistsCheck (OBJ_PO001_SAPGuiButton_Information_Continue, 10)
			If bFlag Then
				kyPerformSAPOperation OBJ_PO001_SAPGuiButton_Information_Continue, "Highlight", ""
				kyPerformSAPOperation OBJ_PO001_SAPGuiButton_Information_Continue, "Click", ""
			End If
		
		kyPerformSAPOperation OBJ_PO001_SAPGuiToolbar_CreateArticleHierarchy_GridToolbar, "PressButton", "NODECREATE"
		kyFillTableRowByRowAndColID OBJ_PO001_SAPGuiGrid_CreateArticleHierarchy_GridViewCtrl, 1, "Hierarchy Node",UNIVERSAL.Item("ChildNode")
		kyFillTableRowByRowAndColID OBJ_PO001_SAPGuiGrid_CreateArticleHierarchy_GridViewCtrl, 1, "Description",UNIVERSAL.Item("ChildNodeDesc")
		
		kyPerformSAPOperation OBJ_PO001_SAPGuiButton_CreateArticleHierarchy_Save, "Click", ""
		kySAPSync OBJ_PO001_SAPGuiStatus_CreateArticleHierarchy_StatusBar, 3
				
				passed = "Execute after filling the hierarchy details"
				failed="Error while executing hierarchy details"
				subWriteTestStepResult passed,failed, true
		
	 End If 
End Sub





