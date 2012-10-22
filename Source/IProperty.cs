using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetComBridge
{
    /// <summary>Interface of the Property class</summary>
    [Guid("efca4c08-529d-437e-bba3-e2a03b6307d6"), ComVisible(true)]
    public interface IProperty
    {
        [Description("Get the property")]
        Instance Get();
        
        [Description("Set the property")]
        void Set(ref object pArgument);
    }
}
