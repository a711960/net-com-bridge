using System;
using System.Runtime.InteropServices;

namespace NetComBridge
{
    /// <summary>
    /// Class representing an instance of class
    /// </summary>
    [Guid("2a01432c-8b2f-480b-b18c-3bd62ed488f0")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Instance : IInstance
    {
        private Bridge lBridge;
        internal System.Type lType = null;
        internal System.Object lInstance = null;
        internal bool lIsReady;
        internal System.String lErrorMessage;

        internal Instance(Bridge pBridge, System.Type pObjectType, System.Object pObject){
            this.lBridge = pBridge;
            this.lType = pObjectType;
            this.lInstance = pObject;
        }

        internal Instance(Bridge pBridge, System.Object pObject){
            this.lBridge = pBridge;
            if(pObject!=null) this.lType = pObject.GetType();
            this.lInstance = pObject;
        }

        internal Instance(Bridge pBridge){
            this.lBridge = pBridge;
        }

        /// <summary>Subscribe to an event</summary>
        /// <param name="pEventName"></param>
        /// <param name="pArguments"></param>
        public void SubscribeEvent(String pEventName, object[] pArguments){
            System.Reflection.EventInfo eventInfo = this.lType.GetEvent(pEventName);
            Delegate handler = Delegate.CreateDelegate(eventInfo.EventHandlerType, this.lBridge, this.lBridge.lHandleEventMethod);
            eventInfo.AddEventHandler(this.lInstance, handler);
        }

        /// <summary>Wait the provided time in millisecond</summary>
        /// <param name="pTimeMillisecond"></param>
        /// <returns></returns>
        public Instance Wait(int pTimeMillisecond){
            System.Threading.Thread.Sleep(pTimeMillisecond);
            return this;
        }

        /// <summary>Get the invocation status for asynchronous call</summary>
        public bool IsReady{
            get { 
                if (this.lErrorMessage != null) throw new ApplicationException(this.lErrorMessage);
                return this.lIsReady;           
            }
        }

        private System.Object CurrentInstance {
            get{
                if (this.lInstance == null) throw new ApplicationException("Instance is null !");
                return this.lInstance;
            }
        }

        /// <summary>Get a property</summary>
        /// <param name="pPropertyName"></param>
        /// <returns></returns>
        public Property Property(string pPropertyName){
            System.Reflection.PropertyInfo lProperty = this.lType.GetProperty(pPropertyName);
            if (lProperty == null) throw new ApplicationException("Property <" + pPropertyName + "> not found! ");
            return new Property(this.lBridge, this.lType, this.CurrentInstance, lProperty);
        }

        /// <summary>Get the Assembly</summary>
        public Assembly Assembly{
            get { return new Assembly( this.lBridge, this.lType.Assembly); }
        }

        /// <summary>Get the Type</summary>
        public Type Type {
            get {return new Type( this.lBridge, this.lType); }
        }

        /// <summary>Convert this object to a boolean</summary>
        public bool ToBool {
            get {return (bool)this.CurrentInstance; }
        }

        /// <summary>Convert this object to an interger</summary>
        public int ToInt {
            get {return (int)this.CurrentInstance; }
        }

        /// <summary>Get the instance value object</summary>
        public System.Object Value {
            get {return lInstance; }
        }

        internal System.Type BaseType {
            get {return this.lType; } 
        }

        /// <summary>Get the list of methods</summary>
        /// <returns></returns>
        public System.String[] GetMethodsList(){
            System.Reflection.MethodInfo[] lMethods = this.lType.GetMethods();
            string[] lRet = new string[lMethods.Length];
            for (int i = 0; i < lMethods.Length; i++){
                string largs=string.Empty;
                System.Reflection.ParameterInfo[] lParameters =lMethods[i].GetParameters();
                for (int j = 0; j < lParameters.Length; j++){
                    if (largs != string.Empty) { largs += ", "; }
                    largs += lParameters[j].ToString();
                }
                lRet[i] = lMethods[i].ReturnType.ToString() + " " + lMethods[i].Name + "(" + largs + ")";
            }
            return lRet;
        }

        /// <summary>Get the list of fields</summary>
        /// <returns></returns>
        public System.String[] GetFieldsList(){
            System.Reflection.FieldInfo[] lFields = this.lType.GetFields();
            string[] lRet = new string[lFields.Length];
            for (int i = 0; i < lFields.Length; i++){
                lRet[i] = lFields[i].FieldType.ToString() + " " + lFields[i].Name;
            }
            return lRet;
        }

        /// <summary>Convert an object to another type</summary>
        /// <param name="pNewType"></param>
        /// <returns></returns>
        public Instance CastAs(string pNewType){
            System.Type lNewType = this.lBridge.GetLoadedTypeFromString(pNewType);
            if (lNewType != null){
                if (lNewType.IsAssignableFrom(this.lType)){
                    return new Instance(this.lBridge, lNewType, this.CurrentInstance);
			    }else{
                    throw new ApplicationException("Cast fail from Type <" + this.lType.FullName + "> to type <" + pNewType + ">! ");
			    }
            }else{
                throw new ApplicationException("Type <" + pNewType + "> not found! ");
            }
        }

        /// <summary>Get a method</summary>
        /// <param name="pMethod"></param>
        /// <returns></returns>
        public Method Method(string pMethod){
           // if (0 != this.lType.GetMember(pMethod, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.InvokeMethod).Length) 
           //       throw new ApplicationException("Method <" + pMethod + "> not found! ");
            return new Method(this.lBridge, this.CurrentInstance, this.lType, pMethod);
        }

        /// <summary>Get a field</summary>
        /// <param name="pFieldName"></param>
        /// <returns></returns>
        public System.Object GetField(string pFieldName){
            try{
                System.Reflection.FieldInfo lField = this.lType.GetField(pFieldName);
                if (lField == null) throw new ApplicationException("Field <" + pFieldName + "> not found! ");
                return lField.GetValue(this.CurrentInstance);
            }catch(ApplicationException e){
                throw new ApplicationException("Field <" + pFieldName + "> failed! \r\n" + e.Message );
            }
        }

        /// <summary>Set a field</summary>
        /// <param name="pFieldName"></param>
        /// <param name="pArgument"></param>
        public void SetField(string pFieldName, ref object pArgument){
            try{
                this.lType.GetField(pFieldName).SetValue(this.CurrentInstance, pArgument);
            }catch(ApplicationException e){
                throw new ApplicationException("Field <" + pFieldName + "> failed! \r\n" + e.Message );
            }
        }

    }
}
