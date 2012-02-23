﻿using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    [Guid("2a01432c-8b2f-480b-b18c-3bd62ed488f0")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Instance : IInstance
    {
        private NetComBridge lBridge;
        public System.Type lType = null;
        public System.Object lInstance = null;
        public bool lIsReady;
        public System.String lErrorMessage;

        internal Instance(NetComBridge pBridge, System.Type pObjectType, System.Object pObject){
            this.lBridge = pBridge;
            this.lType = pObjectType;
            this.lInstance = pObject;
        }

        internal Instance(NetComBridge pBridge, System.Object pObject){
            this.lBridge = pBridge;
            if(pObject!=null) this.lType = pObject.GetType();
            this.lInstance = pObject;
        }

        internal Instance(NetComBridge pBridge){
            this.lBridge = pBridge;
        }

        public Instance Wait(int pTimeMillisecond){
            System.Threading.Thread.Sleep(pTimeMillisecond);
            return this;
        }

        public bool IsReady{
            get { 
                if (this.lErrorMessage != null) throw new System.Exception(this.lErrorMessage);
                return this.lIsReady;           
            }
        }

        private System.Object CurrentInstance{
            get{
                if (this.lInstance == null) throw new System.Exception("Instance is null !");
                return this.lInstance;
            }
        }

        public Property Property(string pPropertyName){
            System.Reflection.PropertyInfo lProperty = this.lType.GetProperty(pPropertyName);
            if (lProperty == null) throw new System.Exception("Property <" + pPropertyName + "> not found! ");
            return new Property(this.lBridge, this.lType, this.CurrentInstance, lProperty);
        }

        public Assembly Assembly{
            get { return new Assembly( this.lBridge, this.lType.Assembly); }
        }

        public Type Type{            
            get {return new Type( this.lBridge, this.lType); }
        }

        public bool ToBool {
            get {return (bool)this.CurrentInstance; }
        }

        public int ToInt {
            get {return (int)this.CurrentInstance; }
        }

        public System.Object Value {
            get {return lInstance; }
        }

        public System.Type BaseType {
            get {return this.lType; } 
        } 

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

        public System.String[] GetFieldsList(){
            System.Reflection.FieldInfo[] lFields = this.lType.GetFields();
            string[] lRet = new string[lFields.Length];
            for (int i = 0; i < lFields.Length; i++){
                lRet[i] = lFields[i].FieldType.ToString() + " " + lFields[i].Name;
            }
            return lRet;
        }

        public Instance CastAs(string pNewType){
            System.Type lNewType = this.lBridge.GetLoadedTypeFromString(pNewType);
            if (lNewType != null){
                if (lNewType.IsAssignableFrom(this.lType)){
                    return new Instance(this.lBridge, lNewType, this.CurrentInstance);
			    }else{
                    throw new System.Exception("Cast fail from Type <" + this.lType.FullName + "> to type <" + pNewType + ">! ");
			    }
            }else{
                throw new System.Exception("Type <" + pNewType + "> not found! ");
            }
        }

        public Method Method(string pMethod){
           // if (0 != this.lType.GetMember(pMethod, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.InvokeMethod).Length) 
           //       throw new System.Exception("Method <" + pMethod + "> not found! ");
            return new Method(this.lBridge, this.CurrentInstance, this.lType, pMethod);
        }

        public System.Object GetField(string pFieldName){
            try{
                System.Reflection.FieldInfo lField = this.lType.GetField(pFieldName);
                if (lField == null) throw new System.Exception("Field <" + pFieldName + "> not found! ");
                return lField.GetValue(this.CurrentInstance);
            }catch(System.Exception e){
                throw new System.Exception("Field <" + pFieldName + "> failed! \r\n" + e.Message );
            }
        }

        public void SetField(string pFieldName, ref object pArgument){
            try{
                this.lType.GetField(pFieldName).SetValue(this.CurrentInstance, pArgument);
            }catch(System.Exception e){
                throw new System.Exception("Field <" + pFieldName + "> failed! \r\n" + e.Message );
            }
        }

    }
}