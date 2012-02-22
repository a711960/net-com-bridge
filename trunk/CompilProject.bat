
@ECHO OFF

REM Requirements :
REM  .NET Framework 3.5
REM  Sandcastle Help File Builder : http://shfb.codeplex.com/

"%SystemRoot%\Microsoft.NET\Framework\v3.5\MSBuild.exe" /p:Configuration=Release /p:SignAssembly=true /p:AssemblyOriginatorKeyFile="NetComBridge.snk" "NetComBridge\NetComBridge.csproj" 
pause
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" /p:CleanIntermediates=True /p:Configuration=Release Documentation.shfbproj
pause
tlbimp comp1.dll /out:NETcomp1.dll
pause
exit