using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    /// <summary>  Class reffering to an assembly's properties </summary>
    [Guid("e5e1eb63-89be-4456-b06c-381a4964b1cb")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Property : NetComBridgeLib.IProperty
    {
        private Bridge lBridge;
        private System.Object lInstance;
        private System.Type lType;
        private System.Reflection.PropertyInfo lProperty;

        internal Property(Bridge pBridge, System.Type pType, object pInstance, System.Reflection.PropertyInfo pProperty)
        {
            this.lBridge = pBridge;
            this.lType = pType;
            this.lInstance = pInstance;
            this.lProperty = pProperty;
        }

        public Instance Get(){
            try{
                object lValue = lProperty.GetValue(this.lInstance, null);
                return new Instance(this.lBridge, lValue);
            }catch (ApplicationException e){
                throw new ApplicationException("Get Property <" + lProperty.Name + "> failed! \r\n" + e.Message);
            }
        }

        public void Set(ref object pArgument){
            try{
                lProperty.SetValue(this.lInstance, pArgument, null);
            }catch (ApplicationException e){
                throw new ApplicationException("Set Property <" + lProperty.Name + "> failed! \r\n" + e.Message);
            }
        }


    }
}
