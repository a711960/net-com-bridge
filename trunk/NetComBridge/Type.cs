using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    [Guid("320814c2-a6cf-4d86-9a7e-ac7f9b68711c")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Type : IType
    {
        private System.Type lType;
        private NetComBridge lBridge;

        public Type(NetComBridge netComBridge, System.Type lType){
            this.lBridge = netComBridge;
            this.lType = lType;
        }

        internal System.Type BaseType{
            get{return this.lType;}
        }

        public Method Method(string pStaticMethodName) {
            //if(0!=this.lType.GetMember(pStaticMethodName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Static).Length)
            //  throw new System.Exception("Static method <" + pStaticMethodName + "> not found! ");
            return new Method(this.lBridge, null, this.lType, pStaticMethodName);
        }

        public Property Property(string pPropertyName){
            try{
                System.Reflection.PropertyInfo lProperty = this.lType.GetProperty(pPropertyName);
                if (lProperty == null) throw new System.Exception("Property <" + pPropertyName + "> not found! ");
                return new Property(this.lBridge, this.lType, null, lProperty);
            }catch (System.Exception e){
                throw new System.Exception("Property <" + pPropertyName + "> failed! \r\n" + e.Message);
            }
        }

        public Instance Field(string pFieldName) {
            try{
                System.Reflection.FieldInfo lField = this.lType.GetField(pFieldName);
                if (lField == null) throw new System.Exception("Field <" + pFieldName + "> not found! ");
                return new Instance(this.lBridge, this.lType, lField.GetValue(null));
            }catch(System.Exception e){
                throw new System.Exception("Field <" + pFieldName + "> failed! \r\n" + e.Message );
            }
        }

        public Instance Instantiate0(){
            return this.Instantiate();
        }

        public Instance Instantiate1(object pArgument1){
            return this.Instantiate(pArgument1);
        }

        public Instance Instantiate2(object pArgument1, object pArgument2){
            return this.Instantiate(pArgument1, pArgument2);
        }

        public Instance Instantiate3(object pArgument1, object pArgument2, object pArgument3){
            return this.Instantiate(pArgument1, pArgument2, pArgument3);
        }

        public Instance Instantiate4(object pArgument1, object pArgument2, object pArgument3, object pArgument4){
            return this.Instantiate(pArgument1, pArgument2, pArgument3, pArgument4);
        }

        public Instance Instantiate5(object pArgument1, object pArgument2, object pArgument3, object pArgument4, object pArgument5){
            return this.Instantiate(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5);
        }

        public Instance Instantiate(params object[] pArguments){
            return this.InstantiateM(ref pArguments);
        }

        public Instance InstantiateM(ref object[] pArguments){
            try{
                System.Object lObject = null;
                if (pArguments.Length==0) {
                    lObject = System.Activator.CreateInstance(this.lType);
                }else{
                    System.Type[] lTypes = new System.Type[pArguments.Length];
                    object[] lArguments = new object[pArguments.Length];
                    for (int i = 0; i < pArguments.Length; i++){
                        if (pArguments[i] is Instance){
                            lTypes[i] = ((Instance)pArguments[i]).BaseType;
                            lArguments[i] = ((Instance)pArguments[i]).Value;
                        }else{
                            lTypes[i] = pArguments[i].GetType();
                            lArguments[i] = pArguments[i];
                        }
                    }
                    System.Reflection.ConstructorInfo  lConstructor = lType.GetConstructor(lTypes);
                    if (lConstructor == null) throw new System.Exception("Contructor <" + this.lType.FullName + "> not found! ");
                    lObject = lConstructor.Invoke(lArguments);
                }
                return new Instance(this.lBridge, lObject);
            }catch (System.Exception e){
                throw new System.Exception("Instantiate failed! \r\n" + e.InnerException.Message);
            }
        }

        public string[] GetMethodsList(){
            System.Reflection.MethodInfo[] lMethods = this.lType.GetMethods();
            System.String[] lRet = new System.String[lMethods.Length];
            for (int i = 0; i < lMethods.Length; i++){
                for (int m = 0; m < lMethods.Length; m++){
                    string largs = string.Empty;
                    System.Reflection.ParameterInfo[] lParameters = lMethods[m].GetParameters();
                    for (int p = 0; p < lParameters.Length; p++){
                        if (largs != string.Empty) { largs += ", "; }
                        largs += lParameters[p].ToString();
                    }
                    lRet[i++] = lMethods[m].ReturnType.FullName + " " + lMethods[m].Name + "(" + largs + ")";
                }
            }
            return lRet;
        }

    }
}
