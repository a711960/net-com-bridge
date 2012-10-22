using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetComBridge
{
    /// <summary>Interface of the Type class</summary>
    [Guid("d45c8a22-df5c-4152-8169-8eb960027624"), ComVisible(true)]
	public interface IType
    {
        [Description("Get a field")]
        Instance Field(string pFieldName);

        [Description("Instanciate a class with the provided arguments")]
        Instance Instantiate([Optional][DefaultParameterValue(null)]object pArgument1, [Optional][DefaultParameterValue(null)]object pArgument2, [Optional][DefaultParameterValue(null)]object pArgument3, [Optional][DefaultParameterValue(null)]object pArgument4, [Optional][DefaultParameterValue(null)]object pArgument5, [Optional][DefaultParameterValue(null)]object pArgument6);

        [Description("Instanciate a class with an arguments array as parameter")]
        Instance InstantiateT(ref object[] pArguments);

        [Description("Get a static method")]
        Method Method(string pStaticMethodName);

        [Description("Get a property")]
        Property Property(string pPropertyName);

        [Description("Get the list of methods")]
        string[] GetMethodsList();
    }
}
