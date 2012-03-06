using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    /// <summary>Interface of the Instance class</summary>
    [Guid("af9a9a70-672f-4e8b-b52a-bee5f9d11b0b"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IInstance
    {
        Instance CastAs(string pNewType);
        object GetField(string pField);
        void SubscribeEvent(String pEventName, object[] arguments);
        string[] GetFieldsList();
        string[] GetMethodsList();
        Method Method(string pMethod);
        Property Property(string pPropertyName);
        void SetField(string pField, ref object pArgument);
        bool ToBool { get; }
        int ToInt { get; }
        object Value { get; }
        Assembly Assembly { get; }
        Type Type { get; }
        bool IsReady { get; }
        Instance Wait(int pTimeMillisecond);
    }
}
