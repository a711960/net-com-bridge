using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetComBridge
{
    /// <summary>Interface of the Assembly class</summary>
    [Guid("e69f585d-259c-4219-bfe8-fc78f6a12b61"), ComVisible(true)]
    public interface IAssembly
    {
        [Description("Returns a list of constructors")]
        string[] GetConstrutorsList();

        [Description("Returns a list of static methods")]
        string[] GetStaticMethodsList();

        [Description("Get the assembly location path")]
        string Location { get; }

        [Description("Returns the assembly name")]
        Type Type(string pFullTypeName);
    }
}
