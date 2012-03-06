using System;
using System.Runtime.InteropServices;

namespace NetCom
{
    /// <summary>Interface of the Method class</summary>
    [Guid("37e0c8ce-1d23-492f-995c-2329125a8b8c"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMethod
    {
        Instance Invoke0();
        Instance Invoke1(object pArgument1);
        Instance Invoke2(object pArgument1, object pArgument2);
        Instance Invoke3(object pArgument1, object pArgument2, object pArgument3);
        Instance Invoke4(object pArgument1, object pArgument2, object pArgument3, object pArgument4);
        Instance Invoke5(object pArgument1, object pArgument2, object pArgument3, object pArgument4, object pArgument5);
        Instance InvokeM(ref object[] pArguments);
        Instance InvokeAsynchrone(ref object[] pArguments);
    }
}
