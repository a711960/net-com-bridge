NetComBridge provides a COM (Component Object Model) bridge that allow you to use .Net libraries within your VBA (Visual Basic for Applications) or VBS (Visual Basic Script) code. The is an equivalent to the DotnetFactory library used in QTP.

This was originally designed for .NET libraries that are not exposed to COM such as [SELENIUM project](http://seleniumhq.org/).

<h4>Release status</h4>
  * Alpha. Feel free to report defects or wanted feature : [Issues](http://code.google.com/p/net-com-bridge/issues/list)

<h4>Features </h4>
  * Load an external library
  * Load an assembly from the GAC
  * Instantiate a Type
  * Invoke a Method (Synchronously or Asynchronously)
  * Get and Set a Property
  * Get and Set a Field
  * Document loaded assemblies, types, methods, properties and fields

<h4>Minimum Requirements </h4>
  * Microsoft Windows XP
  * Microsoft .NET Framework 2

<h4>How to install ? </h4>
  * Make sure the your computer address the minimum requirements
  * Download the bridge installation package : [NetComBridgeSetup-1.0.6](http://code.google.com/p/net-com-bridge/downloads/detail?name=NetComBridgeSetup-1.0.6.exe)
  * Install the package

<h4>How to use in Excel ? </h4>
  * Open an Excel workbook and then open the Visual Basic Editor (Tools > Macro > Visual Basic Editor or ALT+F11).
  * Create a module to write your code (Insert > Module).
  * Add the "NetComBridge Type Library" Reference ( Tools > References... )
  * Insert the following lines of code :
```
Sub Test
    Dim lBridge As New NetComBridge.Bridge
    lBridge.LoadAssembly "System.Windows.Forms", "2.0.0.0", "neutral", "b77a5c561934e089"
    lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invok _
        "My message", _
        "My title", _
        lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("YesNo")
End Sub
```

  * Run the code and enjoy !

<h4>How to use in a vbs script ? </h4>
  * Create a text file.
  * Open the file with notepad and insert the following lines of code :
```
 Set lBridge = CreateObject("NetComBridge.Bridge")
 lBridge.LoadAssembly "System.Windows.Forms", "2.0.0.0", "neutral", "b77a5c561934e089"
 lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invok _
	"My message", _
	"My title", _
	lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("YesNo")
```
  * Save the file with a ".vbs" extension.
  * Execute the file and enjoy !

<h4>How to compile and use CSharp code in Excel ? </h4>
  * Open an Excel workbook and then open the Visual Basic Editor (Tools > Macro > Visual Basic Editor or ALT+F11).
  * Create a module to write your code (Insert > Module).
  * Add the "NetComBridge Type Library" Reference ( Tools > References...)
  * Insert the following lines of code in the module :
```
Public Sub Test_Compiler()
    Dim ass As NetComBridge.Assembly, ret As Integer
    Dim comp As New NetComBridge.Compiler
    
    comp.AddLine "using System;"
    comp.AddLine "namespace MyNameSpace {"
    comp.AddLine "    public class MyClass {"
    comp.AddLine "        public int MyMethod(int x) {"
    comp.AddLine "            return x + 100;"
    comp.AddLine "        }"
    comp.AddLine "    }"
    comp.AddLine "}"
    Set ass = comp.Compile()
    
    ret = ass.Type("MyNameSpace.MyClass").Instantiate().Method("MyMethod").Invok(100).ToInt
    MsgBox "Result is : " & ret
End Sub
```
  * Execute and enjoy !

<h4>Selenium Exemples</h4>
> [Exemples using Selenium libraries](http://code.google.com/p/net-com-bridge/wiki/SeleniumExemples)

<h4>Release note </h4>
  * 1.0.4 - fix installation registration
  * 1.0.3 - first release

<h4>Author</h4>
> Florent BREHERET