using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetComBridge
{
    [Guid("F7FEE81B-DCC7-4f7c-9549-80704E92F062"), ComVisible(true)]
    public interface ICompiler
    {
        [Description("Add a reference library for the compilation.")]
        void addReferenceAssembly(String pAssemblyName);

        [Description("Add a line of code for the compilation")]
        void addLine(String Code);

        [Description("Compile the code and return an Assembly type")]
        NetComBridge.Assembly Compile();
    }
}
