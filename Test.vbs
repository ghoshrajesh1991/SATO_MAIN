Dim App 'As Application
Set App = CreateObject("QuickTest.Application")
App.Launch
App.Visible = False

'Open QTP Test
App.Open "C:\SATO_MAIN\Layer_Test\TestScripts\SATO", TRUE 'Set the QTP test path
App.Test.Run

App.Test.Close

App.Quit