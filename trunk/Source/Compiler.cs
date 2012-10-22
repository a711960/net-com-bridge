using System;
using System.Runtime.InteropServices;
using System.Text;
using System.ComponentModel;

namespace NetComBridge
{
    /// <summary>Class to compile CSharp code and return an Assembly Type</summary>
    /// <example>
    /// This example compile the calculation class, call the method and show the result.
    /// <code lang="vbs">	
    /// Dim comp, ass, ret
    /// Set comp = CreateObject("NetComBridge.Compiler")
    /// 
    /// comp.addLine "using System;"
    /// comp.addLine "namespace MyNameSpace {"
    /// comp.addLine "    public class MyClass {"
    /// comp.addLine "        public int MyMethod(int x) {"
    /// comp.addLine "            return x + 100;"
    /// comp.addLine "        }"
    /// comp.addLine "    }"
    /// comp.addLine "}"
    /// Set ass = comp.Compile()
    ///
    /// ret = ass.Type("MyNameSpace.MyClass").Instantiate().Method("MyMethod").Invok(100).ToInt
    /// MsgBox "Result is : " &amp; ret
    /// </code>
    /// </example>
    /// 

    [Description("Class to compile CSharp code and return an Assembly Type"), ProgId("NetComBridge.Compiler")]
    [Guid("00CEE59E-6A4B-4cf2-8AD3-E6261759CA24"), ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Compiler : ICompiler
    {
        private System.CodeDom.Compiler.CodeDomProvider provider;
        private System.CodeDom.Compiler.CompilerParameters cpParams;
        private System.Text.StringBuilder sbCode;
        static NetComBridge.Assembly assembly;
        static StringBuilder sbPreviousCode;
        static bool isAlreadyCompiled;

        public Compiler() {
            this.provider = System.CodeDom.Compiler.CodeDomProvider.CreateProvider("CSharp");
            this.cpParams = new System.CodeDom.Compiler.CompilerParameters();
            this.cpParams.GenerateInMemory = false;
            this.cpParams.GenerateExecutable = false;
            //   this.cpParams.OutputAssembly = @"C:\" + DateTime.Now.ToString("Mdyyyy") + ".dll";
            this.sbCode = new StringBuilder();
        }

        /// <summary>Add a reference library for the compilation.</summary>
        /// <param name="pAssemblyName"></param>
        public void addReferenceAssembly(String pAssemblyName){
            this.cpParams.ReferencedAssemblies.Add(pAssemblyName);
            //this.sbHeader.Append("using " + System.Reflection.Assembly.LoadFrom(pAssemblyName).GetName().Name + ";" + Environment.NewLine);
        }

        /// <summary>Add a line of code for the compilation</summary>
        /// <param name="LineOfCode"></param>
        public void addLine(String LineOfCode){
            this.sbCode.AppendLine(LineOfCode);
        }

        /// <summary>Compile the code and return an Assembly type</summary>
        /// <returns></returns>
        public NetComBridge.Assembly Compile(){
            if ( !Compiler.isAlreadyCompiled) {
                System.CodeDom.Compiler.CompilerResults crResults = this.provider.CompileAssemblyFromSource(this.cpParams, this.sbCode.ToString());
                if (crResults.Errors.HasErrors){
                    StringBuilder sbErrors = new StringBuilder();
                    foreach (System.CodeDom.Compiler.CompilerError err in crResults.Errors) {
                        sbErrors.AppendLine(System.Text.RegularExpressions.Regex.Replace( err.ToString(), @"\w:\\[^(]+", "line") );
                    }
                    throw new System.ApplicationException(sbErrors.ToString() );
                }else{
                    System.Reflection.Assembly compiledAssembly = crResults.CompiledAssembly;
                    NetComBridge.Bridge bridge = new Bridge();
                    Compiler.assembly = new Assembly(bridge, compiledAssembly);
                    Compiler.isAlreadyCompiled = true;
                    Compiler.sbPreviousCode = this.sbCode;
                    return assembly;
                }
            }else if (!Compiler.sbPreviousCode.Equals(this.sbCode)) {
                throw new ApplicationException("The code is already compiled.\r\nReopen Excel to compile again.\r\n");
            }else{
                NetComBridge.Bridge bridge = new Bridge();
                return Compiler.assembly;
            }
        }
    }
}