' Quick test with System.Windows.Forms assembly
' Author: Florent BREHERET
' Description: This example compile the calculation class, call the method and show the result.

Dim comp, ass, ret

Set comp = CreateObject("NetComBridge.Compiler")

comp.addLine "using System;"
comp.addLine "namespace MyNameSpace {"
comp.addLine "    public class MyClass {"
comp.addLine "        public int MyMethod(int x) {"
comp.addLine "            return x + 100;"
comp.addLine "        }"
comp.addLine "    }"
comp.addLine "}"
Set ass = comp.Compile()

ret = ass.Type("MyNameSpace.MyClass").Instantiate().Method("MyMethod").Invok(100).ToInt
MsgBox "Result is : " & ret