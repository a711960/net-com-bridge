using System;
using System.Runtime.InteropServices;

namespace NetCom
{
    /// <summary>Interface of the Assembly class</summary>
    [Guid("e69f585d-259c-4219-bfe8-fc78f6a12b61"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAssembly
    {
        string[] GetConstrutorsList();
        string[] GetStaticMethodsList();
        string Location { get; }
        Type Type(string pFullTypeName);
    }
}
