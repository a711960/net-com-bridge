using System;
using System.Runtime.InteropServices;

namespace NetComBridge
{
    /// <summary>
    /// Class reffering to types
    /// </summary>
    [Guid("320814c2-a6cf-4d86-9a7e-ac7f9b68711c")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Type : IType
    {
        private System.Type lType;
        private Bridge lBridge;

        public Type(Bridge pBridge, System.Type pType){
            this.lBridge = pBridge;
            this.lType = pType;
        }

        internal System.Type BaseType{
            get{return this.lType;}
        }

        /// <summary>Get a static method</summary>
        /// <param name="pStaticMethodName"></param>
        /// <returns></returns>
        public Method Method(string pStaticMethodName) {
            //if(0!=this.lType.GetMember(pStaticMethodName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Static).Length)
            //  throw new ApplicationException("Static method <" + pStaticMethodName + "> not found! ");
            return new Method(this.lBridge, null, this.lType, pStaticMethodName);
        }

        /// <summary>Get a property</summary>
        /// <param name="pPropertyName"></param>
        /// <returns></returns>
        public Property Property(string pPropertyName){
            try{
                System.Reflection.PropertyInfo lProperty = this.lType.GetProperty(pPropertyName);
                if (lProperty == null) throw new ApplicationException("Property <" + pPropertyName + "> not found! ");
                return new Property(this.lBridge, this.lType, null, lProperty);
            }catch (ApplicationException e){
                throw new ApplicationException("Property <" + pPropertyName + "> failed! \r\n" + e.Message);
            }
        }

        /// <summary>Get a field</summary>
        /// <param name="pFieldName"></param>
        /// <returns></returns>
        public Instance Field(string pFieldName) {
            try{
                System.Reflection.FieldInfo lField = this.lType.GetField(pFieldName);
                if (lField == null) throw new ApplicationException("Field <" + pFieldName + "> not found! ");
                return new Instance(this.lBridge, this.lType, lField.GetValue(null));
            }catch(ApplicationException e){
                throw new ApplicationException("Field <" + pFieldName + "> failed! \r\n" + e.Message );
            }
        }

        /// <summary>Instanciate a class with the provided arguments</summary>
        /// <param name="pArgument1"></param>
        /// <param name="pArgument2"></param>
        /// <param name="pArgument3"></param>
        /// <param name="pArgument4"></param>
        /// <param name="pArgument5"></param>
        /// <param name="pArgument6"></param>
        /// <returns></returns>
        public Instance Instantiate([Optional][DefaultParameterValue(null)]object pArgument1, [Optional][DefaultParameterValue(null)]object pArgument2, 
            [Optional][DefaultParameterValue(null)]object pArgument3, [Optional][DefaultParameterValue(null)]object pArgument4,
            [Optional][DefaultParameterValue(null)]object pArgument5, [Optional][DefaultParameterValue(null)]object pArgument6){

            if(object.Equals(pArgument1,null)){
                return this.InstantiateParams();
            }else if (object.Equals(pArgument2, null)){
                return this.InstantiateParams(pArgument1);
            }else if (object.Equals(pArgument3, null)){
                return this.InstantiateParams(pArgument1, pArgument2);
            }else if (object.Equals(pArgument4, null)){
                return this.InstantiateParams(pArgument1, pArgument2, pArgument3);
            }else if (object.Equals(pArgument5, null)){
                return this.InstantiateParams(pArgument1, pArgument2, pArgument3, pArgument4);
            }else if (object.Equals(pArgument6, null)){
                return this.InstantiateParams(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5);
            }else{
                return this.InstantiateParams(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5, pArgument6);
            }
        }

        private Instance InstantiateParams(params object[] pArguments){
            return this.InstantiateT(ref pArguments);
        }

        /// <summary>Instanciate a class with an arguments array as parameter</summary>
        /// <param name="pArguments"></param>
        /// <returns></returns>
        public Instance InstantiateT(ref object[] pArguments){
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
                    if (lConstructor == null) throw new ApplicationException("Contructor <" + this.lType.FullName + "> not found! ");
                    lObject = lConstructor.Invoke(lArguments);
                }
                return new Instance(this.lBridge, lObject);
            }catch (ApplicationException e){
                throw new ApplicationException("Instantiate failed! \r\n" + e.InnerException.Message);
            }
        }

        /// <summary>Get the list of methods</summary>
        /// <returns></returns>
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
