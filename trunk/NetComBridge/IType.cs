using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    /// <summary>Interface of the Type class</summary>
    [Guid("d45c8a22-df5c-4152-8169-8eb960027624"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IType
    {
        Instance Field(string pFieldName);
        Instance Instantiate0();
        Instance Instantiate1(object pArgument1);
        Instance Instantiate2(object pArgument1, object pArgument2);
        Instance Instantiate3(object pArgument1, object pArgument2, object pArgument3);
        Instance Instantiate4(object pArgument1, object pArgument2, object pArgument3, object pArgument4);
        Instance Instantiate5(object pArgument1, object pArgument2, object pArgument3, object pArgument4, object pArgument5);
        Instance InstantiateM(ref object[] pArguments);
        Method Method(string pStaticMethodName);
        Property Property(string pPropertyName);
        string[] GetMethodsList();
    }
}
