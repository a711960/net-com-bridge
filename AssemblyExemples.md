### Show the current date in a message box ###

```
 Sub Test_DateTime()
     Dim lBridge As New NetComBridge.Bridge
     MsgBox lBridge.Type("System.DateTime").Property("Now").Get().Method("ToString").Invok("yyyy MM dd")
 End Sub
```

### Use a .NET form object ###
```
Sub Test_MessageBox
    Dim lBridge As New NetComBridge.Bridge
    lBridge.LoadAssembly "System.Windows.Forms", "2.0.0.0", "neutral", "b77a5c561934e089"
    lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invok _
        "My message", _
        "My title", _
        lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("YesNo")
End Sub
```