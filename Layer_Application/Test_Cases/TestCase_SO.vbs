

Sub TC_SO_TC001_Create_Sales_Order()
	
	sub_LaunchSAPApplication
	sub_SelectServer
	sub_UserLogin
	sub_PageNavigate "/nva01", OBJ_SO001_SAPGuiEdit_CreateSales_OrderType, "Create Sales Document"
	sub_createSalesOrder
	sub_SalesOrderConfirmBeforSave
	sub_SalesOrderNumberFetch
	sub_PageNavigate "/nva03", OBJ_SO001_SAPGuiEdit_DisplaySalesDoc_Order, "Display Sales Document"
	sub_SalesOrderNumberDisplay UNIVERSAL.Item("SO_Number")
	sub_ToolBarMenuNavigate "Goto;Header;Status"
	sub_SalesOrderStatusVerify OBJ_SO001_SAPGuiEdit_CreateSales_OverallCredStat, "OverallCredStat", UNIVERSAL.Item("OverallCredStat")
'	sub_SAPLogOff
'	sub_CloseSAPERP
	
End Sub

