using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    [Guid("268c5c67-5f85-498b-a2b1-9fff86495cce"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INetComBridge
    {
        Assembly Assembly(string pAssemblyName);
        string[] GetAssembliesList();
        string[] GetTypesList();
        Assembly LoadAssembly(string pAssemblyName, string pVersion, string pCulture, string pPublicKeyToken);
        Assembly LoadLibrary(string pDllPath);
        Type Type(string pFullTypeName);
    }
}
