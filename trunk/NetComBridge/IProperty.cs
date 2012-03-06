using System;
using System.Runtime.InteropServices;

namespace NetCom
{
    /// <summary>Interface of the Property class</summary>
	[Guid("efca4c08-529d-437e-bba3-e2a03b6307d6"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IProperty
    {
        Instance Get();
        void Set(ref object pArgument);
    }
}
