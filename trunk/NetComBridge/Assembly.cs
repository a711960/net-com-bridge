using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    /// <summary>  Class reffering to an assembly </summary>
    [Guid("c8916909-418b-4616-900f-e76e6f7369ac")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Assembly : NetComBridgeLib.IAssembly
    {
        private System.Reflection.Assembly lAssembly=null;
        private Bridge lBridge=null;

        internal Assembly(Bridge netComBridge, System.Reflection.Assembly pAssembly){
            this.lBridge = netComBridge;
            this.lAssembly = pAssembly;
        }


        /// <summary>Returns the assembly's location</summary>
        /// <returns>path of this assembly</returns>
        public System.String Location{
            get { return this.lAssembly.Location; }
        }

        /// <summary>Returns the type object</summary>
        /// <param name="pFullTypeName">Type name</param>
        /// <returns>Type object</returns>
        public Type Type(string pFullTypeName){
            System.Type lType;
            lType = this.lBridge.GetLoadedTypeFromString(pFullTypeName);
            if (lType!=null){
                return new Type(this.lBridge , lType);
            }else{
                throw new ApplicationException("Type <" + pFullTypeName + "> not found! ");
            }
        }

        /// <summary>Returns the assembly's constructors</summary>
        /// <returns>Returns a string array</returns>
        public System.String[] GetConstrutorsList(){
            int i=0;
            string[] lRet = new string[200];
            System.Type[] lTypes = this.lAssembly.GetTypes( );
            for (int t = 0; t < lTypes.Length; t++){
                if(lTypes[t].IsPublic){
                    System.Reflection.ConstructorInfo[] lConstructors = lTypes[t].GetConstructors(System.Reflection.BindingFlags.Public);
                    for (int c = 0; c < lConstructors.Length; c++){
                        string largs = string.Empty;
                        System.Reflection.ParameterInfo[] lParameters = lConstructors[c].GetParameters();
                        for (int p = 0; p < lParameters.Length; p++)
                        {
                            if (largs != string.Empty) { largs += ", "; }
                            largs += lParameters[p].ToString();
                        }
                        lRet[i++] = lTypes[t].FullName + "(" + largs + ")";
                    }
                }
            }
            return lRet;
        }

        /// <summary>Returns the assembly's static methods</summary>
        /// <returns>Returns a string array</returns>
        public System.String[] GetStaticMethodsList(){
            int i=0;
            string[] lRet = new string[200];
            System.Type[] lTypes = this.lAssembly.GetTypes();
            for (int t = 0; t < lTypes.Length; t++){
                if(lTypes[t].IsPublic){
                    System.Reflection.MethodInfo[] lMethods = lTypes[t].GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    for (int m = 0; m < lMethods.Length; m++){
                        string largs = string.Empty;
                        System.Reflection.ParameterInfo[] lParameters = lMethods[m].GetParameters();
                        for (int p = 0; p < lParameters.Length; p++){
                            if (largs != string.Empty) { largs += ", "; }
                            largs += lParameters[p].ToString();
                        }
                        lRet[i++] = lTypes[t].FullName + "." + lMethods[m].Name + "(" + largs + ")";
                    }
                }
            }
            return lRet;
        }

    }
}
