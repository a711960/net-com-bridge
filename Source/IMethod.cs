using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetComBridge
{
    /// <summary>Interface of the Method class</summary>
    [Guid("37e0c8ce-1d23-492f-995c-2329125a8b8c"), ComVisible(true)]
    public interface IMethod
    {
        [Description("Invoke the method and wait the end of execution")]
        Instance Invok([Optional][DefaultParameterValue(null)]object pArgument1, [Optional][DefaultParameterValue(null)]object pArgument2, [Optional][DefaultParameterValue(null)]object pArgument3, [Optional][DefaultParameterValue(null)]object pArgument4, [Optional][DefaultParameterValue(null)]object pArgument5, [Optional][DefaultParameterValue(null)]object pArgument6);

        [Description("Same as \"Invok\" but with an arguments array as parameter")]
        Instance InvokT(ref object[] pArguments);

        [Description("Invoke the method but don't wait the end of execution")]
        Instance InvokAsync([Optional][DefaultParameterValue(null)]object pArgument1, [Optional][DefaultParameterValue(null)]object pArgument2, [Optional][DefaultParameterValue(null)]object pArgument3, [Optional][DefaultParameterValue(null)]object pArgument4, [Optional][DefaultParameterValue(null)]object pArgument5, [Optional][DefaultParameterValue(null)]object pArgument6);

        [Description("Same as \"InvokAsync\" but with an arguments array as parameter")]
        Instance InvokAsyncT(ref object[] pArguments);
    }
}
