using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    /// <summary>
    /// Class reffering to the bridge
    /// </summary>
    [Guid("d43f6126-5995-4e81-aec9-4b58e66551d9")]
    [ClassInterface(ClassInterfaceType.None)]
    public class NetComBridge : NetComBridgeLib.INetComBridge
    {
        System.Collections.Generic.Dictionary<string, System.Type> cTypes;
        System.Collections.Generic.Dictionary<string, System.Reflection.Assembly> cAssemblies;
        System.AppDomain lDomain;
        int lTimeout;

        public NetComBridge(){
            this.lDomain = System.AppDomain.CurrentDomain;
            this.LoadTypes();
            this.lTimeout = 30000;
        }

        public void Wait(int pTimeMillisecond){
            System.Threading.Thread.Sleep(pTimeMillisecond);
        }

        public int Timeout{
            get{return this.lTimeout;}
            set{this.lTimeout=value;}
        }

        public Type Type(string pFullTypeName){
            System.Type lType;
            lType = this.GetLoadedTypeFromString(pFullTypeName);
            if (lType!=null){
                return new Type(this , lType);
            }else{
                throw new ApplicationException("Type <" + pFullTypeName + "> not found! ");
            }
        }

        public Assembly Assembly(string pAssemblyName){
            System.Reflection.Assembly lAssembly;
            if (this.cAssemblies.TryGetValue(pAssemblyName, out lAssembly)){
                return new Assembly(this, lAssembly);
            }else{
                throw new ApplicationException("Assembly <" + pAssemblyName + "> not found! ");
            }
        }

        internal System.Type GetLoadedTypeFromString(string pFullTypeName){
            System.Type lType;
            if (this.cTypes.TryGetValue(pFullTypeName, out lType)){
                return lType;
            }else{
                return null;
            }
        }

        private void LoadTypes(){
            if (this.cTypes!=null) this.cTypes.Clear();
            this.cTypes = new System.Collections.Generic.Dictionary<System.String, System.Type>();
            this.cAssemblies = new System.Collections.Generic.Dictionary<System.String, System.Reflection.Assembly>();
            System.Reflection.Assembly[] lAssemblies = this.lDomain.GetAssemblies();
            string[] lRet = new string[lAssemblies.Length];
            for (int a = 0; a < lAssemblies.Length; a++){
                this.cAssemblies.Add(lAssemblies[a].FullName, lAssemblies[a]);
                System.Type[] lTypes = lAssemblies[a].GetExportedTypes();
                for (int t = 0; t < lTypes.Length; t++){
                    try { 
                       this.cTypes.Add(lTypes[t].FullName, lTypes[t]); 
                    }catch(ApplicationException e){
                       //System.Console.WriteLine(lTypes[t].FullName);
                    }
                }
            }
        }

        public Assembly LoadLibrary(string pDllPath){
            if (! System.IO.File.Exists(pDllPath)) throw new ApplicationException("Library not found : " + pDllPath);
            System.Reflection.Assembly lAssembly;
            try{
                lAssembly = System.Reflection.Assembly.LoadFrom(pDllPath);
            }catch (ApplicationException e){
                throw new ApplicationException("Failed to load the library : " + pDllPath + "\r\n" + e.Message);
            }
            this.LoadTypes();
            return new Assembly(this, lAssembly);
        }

        public Assembly LoadAssembly(string pAssemblyName, string pVersion, string pCulture, string pPublicKeyToken){
            try{
                string lName = pAssemblyName + ", Version=" + pVersion + ", Culture=" + pCulture + ", PublicKeyToken=" + pPublicKeyToken;
                System.Reflection.Assembly lAssembly = this.lDomain.Load(lName);
                if (lAssembly == null){
                    throw new ApplicationException("Failed to load the library : " + pAssemblyName );
                }else{
                    this.LoadTypes();
                    return new Assembly(this, lAssembly);
                }
            }catch (ApplicationException e){
                throw new ApplicationException("Failed to load the library : " + pAssemblyName + "\r\n" + e.Message);
            }
        }

        public System.String[] GetTypesList(){
            string[] lRet = new string[cTypes.Count];
            cTypes.Keys.CopyTo(lRet, 0);
            return lRet;
        }

        public System.String[] GetAssembliesList(){
            string[] lRet = new string[cAssemblies.Count];
            cAssemblies.Keys.CopyTo(lRet, 0);
            return lRet;
        }

    }
}
