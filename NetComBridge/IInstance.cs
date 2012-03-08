using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetComBridge
{
    /// <summary>Interface of the Instance class</summary>
    [Guid("af9a9a70-672f-4e8b-b52a-bee5f9d11b0b"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IInstance
    {
        [Description("Convert an object to another type")]
        Instance CastAs(string pNewType);

        [Description("Get a field")]
        object GetField(string pField);

        [Description("Subscribe to an event")]
        void SubscribeEvent(String pEventName, object[] arguments);

        [Description("Get the list of fields")]
        string[] GetFieldsList();

        [Description("Get the list of methods")]
        string[] GetMethodsList();

        [Description("Get a method")]
        Method Method(string pMethod);

        [Description("Get a property")]
        Property Property(string pPropertyName);

        [Description("Set a field")]
        void SetField(string pField, ref object pArgument);

        [Description("Convert this object to a boolean")]
        bool ToBool { get; }

        [Description("Convert this object to an interger")]
        int ToInt { get; }

        [Description("Get the instance value object")]
        object Value { get; }

        [Description("Get the Assembly")]
        Assembly Assembly { get; }

        [Description("Get the Type")]
        Type Type { get; }

        [Description("Get the invocation status for asynchronous call")]
        bool IsReady { get; }

        [Description("Wait the provided time in millisecond")]
        Instance Wait(int pTimeMillisecond);
    }
}
