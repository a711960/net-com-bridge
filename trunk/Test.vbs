

 Set lBridge = CreateObject("NetComBridgeLib.NetComBridge")
  
 lBridge.LoadAssembly "System.Windows.Forms", "2.0.0.0", "neutral", "b77a5c561934e089"
 lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invoke3 _
        "My message", _
        "My title", _
        lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("YesNo")