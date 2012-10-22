

Dim lBridge, selenium
Set lBridge = CreateObject("NetComBridge.Bridge")

lBridge.LoadLibrary "C:\net35\WebDriver.dll"
lBridge.LoadLibrary "C:\net35\Selenium.WebDriverBackedSelenium.dll"
lBridge.LoadLibrary "C:\net35\ThoughtWorks.Selenium.Core.dll"
Set selenium = lBridge.Type("Selenium.WebDriverBackedSelenium").Instantiate( _
		lBridge.Type("OpenQA.Selenium.Chrome.ChromeDriver").Instantiate(), _
		"http://www.google.com" _
	).CastAs("Selenium.ISelenium")
selenium.Method("Start").Invok
selenium.Method("Open").Invok "http://www.google.com"
selenium.Method("Type").Invok "name=q", "England"
selenium.Method("Click").Invok "name=btnG"
Debug.Print selenium.Method("GetTitle").Invok()
selenium.Method("Close").Invok
