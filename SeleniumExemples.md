### Use Selenium WebDriver library to drive Firefox ###
The following code launch Firefox and search the word "England"

```
Sub Test_WebDriver()
    Dim lBridge As New NetComBridge.Bridge
    Dim webdriver As NetComBridge.Instance
    lBridge.LoadLibrary "C:\net35\WebDriver.dll"
    Set webdriver = lBridge.Type("OpenQA.Selenium.Firefox.FirefoxDriver").Instantiate().CastAs("OpenQA.Selenium.IWebDriver")
    webdriver.Method("Get").Invok "http://www.google.com"
    webdriver.Method("FindElement").Invok( _
            lBridge.Type("OpenQA.Selenium.By").Method("Name").Invok("q") _
        ).Method("SendKeys").Invok ("England")
    webdriver.Method("FindElement").Invok( _
            lBridge.Type("OpenQA.Selenium.By").Method("Name").Invok("btnG") _
        ).Method("Click").Invok
    webdriver.Method("Quit").Invok
End Sub
```

### Use Selenium WebDriverBackedSelenium library to drive Firefox ###
The following code launch Firefox and search the word "England"
```
Sub Test_WebDriverBackedSelenium()
    Dim lBridge As New NetComBridge.Bridge
    Dim selenium As NetComBridge.Instance
    lBridge.LoadLibrary "C:\net35\WebDriver.dll"
    lBridge.LoadLibrary "C:\net35\Selenium.WebDriverBackedSelenium.dll"
    lBridge.LoadLibrary "C:\net35\ThoughtWorks.Selenium.Core.dll"
    Set selenium = lBridge.Type("Selenium.WebDriverBackedSelenium").Instantiate( _
            lBridge.Type("OpenQA.Selenium.Firefox.FirefoxDriver").Instantiate(), _
            "http://www.google.com" _
        ).CastAs("Selenium.ISelenium")
    selenium.Method("Start").Invok
    selenium.Method("Open").Invok "/"
    selenium.Method("Type").Invok "name=q", "England"
    selenium.Method("Click").Invok "name=btnG"
    selenium.Method("Close").Invok
End Sub
```