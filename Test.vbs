Dim App 'As Application
Set App = CreateObject("QuickTest.Application")
App.Launch
App.Visible = False
' App.Test.Settings.Launchers("SAP").Active = True
' App.Test.Settings.Launchers("SAP").Server = ""
' App.Test.Settings.Launchers("SAP").AutoLogon = False
' App.Test.Settings.Launchers("SAP").User = ""
' App.Test.Settings.Launchers("SAP").Client = ""
' App.Test.Settings.Launchers("SAP").Password = ""
' App.Test.Settings.Launchers("SAP").Language = ""
' App.Test.Settings.Launchers("SAP").RememberPassword = False
' App.Test.Settings.Launchers("SAP").CloseOnExit = False
' App.Test.Settings.Launchers("SAP").IgnoreExistingSessions = True
' App.Test.Settings.Launchers("Web").Active = False
' App.Test.Settings.Launchers("Web").Browser = "IE"
' App.Test.Settings.Launchers("Web").Address = "http://newtours.demoaut.com "
' App.Test.Settings.Launchers("Web").CloseOnExit = True
' App.Test.Settings.Launchers("Windows Applications").Active = True
' App.Test.Settings.Launchers("Windows Applications").Applications.RemoveAll
' App.Test.Settings.Launchers("Windows Applications").RecordOnQTDescendants = True
' App.Test.Settings.Launchers("Windows Applications").RecordOnExplorerDescendants = False
' App.Test.Settings.Launchers("Windows Applications").RecordOnSpecifiedApplications = True
' App.Test.Settings.Run.IterationMode = "rngAll"
' App.Test.Settings.Run.StartIteration = 1
' App.Test.Settings.Run.EndIteration = 1
' App.Test.Settings.Run.ObjectSyncTimeOut = 20000
' App.Test.Settings.Run.DisableSmartIdentification = False
' App.Test.Settings.Run.OnError = "Dialog"
' App.Test.Settings.Resources.DataTablePath = "<Default>"
' App.Test.Settings.Resources.Libraries.RemoveAll
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\All_Templates.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_Constants.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_Database.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_DateTime.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_DB.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_Excel.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_Execution.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_File.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_KeywordsSAP.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_KeywordsWin.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_MethodsGeneral.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\lib_String.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Application\Business_Functions\BF_Common.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Application\Business_Functions\BF_PO.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Application\OR\OR_PO.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Application\Test_Cases\TestCase_PO.vbs")
' App.Test.Settings.Resources.Libraries.Add("C:\Automations\SAPDemo\Layer_Platform\Executioner.vbs")
' App.Test.Settings.Web.BrowserNavigationTimeout = 60000
' App.Test.Settings.Web.ActiveScreenAccess.UserName = ""
' App.Test.Settings.Web.ActiveScreenAccess.Password = ""
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' System Local Monitoring settings
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' App.Test.Settings.LocalSystemMonitor.Enable = false
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' Log Tracking settings
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' With App.Test.Settings.LogTracking 
	' .IncludeInResults = False 
	' .Port = 18081 
	' .IP = "127.0.0.1" 
	' .MinTriggerLevel = "ERROR" 
	' .EnableAutoConfig = False 
	' .RecoverConfigAfterRun = False 
	' .ConfigFile = "" 
	' .MinConfigLevel = "WARN" 
' End With

'Open QTP Test
App.Open "C:\SATO_MAIN\Layer_Test\TestScripts\SATO", TRUE 'Set the QTP test path
App.Test.Run

App.Test.Close

App.Quit