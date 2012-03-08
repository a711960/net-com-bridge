using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetComBridge
{
    [Guid("ed5392cc-fdfb-4cb9-8370-4f4487c18b76")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEvents
    {
        [Description("Events subscribed with \"instance.SubscribeEvent( pEventName ) \" are raised here")]
        void Event(String name, Object sender, Object arguments);
    }

    /// <summary>Interface of the NetComBridge class</summary>
    [Guid("268c5c67-5f85-498b-a2b1-9fff86495cce"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBridge
    {
        [Description("Get an assembly by name")]
        Assembly Assembly(string pAssemblyName);

        [Description("Get the list of assemblies")]
        string[] GetAssembliesList();

        [Description("Get the list of types")]
        string[] GetTypesList();

        [Description("Load an assembly from the GAC (C:\\WINDOWS\\assembly)")]
        Assembly LoadAssembly(string pAssemblyName, string pVersion, string pCulture, string pPublicKeyToken);

        [Description("Load a dll by path")]
        Assembly LoadLibrary(string pDllPath);

        [Description("Get a Type")]
        Type Type(string pFullTypeName);

        [Description("Waits the provided time (millisecond)")]
        void Wait(int pTimeMillisecond);

        [Description("Get/Set the Timeout for asynchronous call")]
        int Timeout { get; set; }
    }
}
