
Sub TC_PO_TC0001_Create_Article_Node()

	sub_LaunchSAPApplication
	sub_SelectServer
	sub_UserLogin
	Sub_CreateArticleNode 
	Sub_HierachySaveVerify
	sub_PageNavigate "/nwmatgrp02", OBJ_PO001_SAPGuiButton_CreateArticleHierarchy_HierarchyData, "Change Article Hierarchy"
	Sub_ModifyArticleNode
	sub_SAPLogOff
	sub_CloseSAPERP
	
End Sub



