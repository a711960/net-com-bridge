
SET ROOT=%CD%
SET PATH=C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\

cd NetComBridge\bin\Release

regasm /u NetComBridgeLib.dll /tlb:NetComBridgeLib.tlb
pause
regasm NetComBridgeLib.dll /tlb:NetComBridgeLib.tlb /codebase

pause
exit
