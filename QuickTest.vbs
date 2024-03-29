' Quick test with System.Windows.Forms assembly
' Author: Florent BREHERET
' Description: Ask the user to show the current date

Dim lBridge, ret

Set lBridge = CreateObject("NetComBridge.Bridge")

lBridge.LoadAssembly "System.Windows.Forms", "2.0.0.0", "neutral", "b77a5c561934e089"
ret=lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invok( _
	"Do you want to know the current date ? ", _
	"NetComBridge Test", _
	lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("YesNo") _
)

If ret=lBridge.Type("System.Windows.Forms.DialogResult").Field("Yes") Then
	lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invok _
		lBridge.Type("System.DateTime").Property("Now").Get().Method("ToString").Invok("dddd dd MMMM yyyy"), _
		"Current date", _
		lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("OK")
End If