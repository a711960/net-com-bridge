using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    [Guid("ed5392cc-fdfb-4cb9-8370-4f4487c18b76")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEvents
    {
        void Event(String name, Object sender, Object arguments);
    }

    /// <summary>Interface of the NetComBridge class</summary>
    [Guid("268c5c67-5f85-498b-a2b1-9fff86495cce"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBridge
    {
        Assembly Assembly(string pAssemblyName);
        string[] GetAssembliesList();
        string[] GetTypesList();
        Assembly LoadAssembly(string pAssemblyName, string pVersion, string pCulture, string pPublicKeyToken);
        Assembly LoadLibrary(string pDllPath);
        Type Type(string pFullTypeName);
    }
}
