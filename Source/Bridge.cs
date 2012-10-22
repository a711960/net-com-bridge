using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.ComponentModel;

namespace NetComBridge
{
    /// <summary>Class to instanciate and use object members</summary>
    /// <example>
    /// 
    /// This example asks the user to show the current date.
    /// <code lang="vbs">	
    /// Dim lBridge, ret
    /// 
    /// Set lBridge = CreateObject("NetComBridge.Bridge")
    /// 
    /// lBridge.LoadAssembly "System.Windows.Forms", "2.0.0.0", "neutral", "b77a5c561934e089"
    /// ret=lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invok( _
    /// 	"Do you want to know the current date ? ", _
    /// 	"NetComBridge Test", _
    /// 	lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("YesNo") _
    /// )
    /// 
    /// If ret=lBridge.Type("System.Windows.Forms.DialogResult").Field("Yes") Then
    /// 	lBridge.Type("System.Windows.Forms.MessageBox").Method("Show").Invok _
    /// 		lBridge.Type("System.DateTime").Property("Now").Get().Method("ToString").Invok("dddd dd MMMM yyyy"), _
    /// 		"Current date", _
    /// 		lBridge.Type("System.Windows.Forms.MessageBoxButtons").Field("OK")
    /// End If
    /// </code>
    /// </example>
    /// 

    [Description("Class to instanciate and use object members"), ProgId("NetComBridge.Bridge")]
    [Guid("d43f6126-5995-4e81-aec9-4b58e66551d9"), ComVisible(true), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IEvents))]
    public class Bridge :  IBridge
    {

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern short GetKeyState(int virtualKeyCode);

        System.Collections.Generic.Dictionary<string, System.Type> cTypes;
        System.Collections.Generic.Dictionary<string, System.Reflection.Assembly> cAssemblies;
        System.AppDomain lDomain;
        int lTimeout;
        internal System.Timers.Timer timerhotkey;
        internal System.Threading.Thread thread;

        [ComVisible(false)]
        public delegate void EventHandler(String name, Object sender, Object arguments); 
        public System.Reflection.MethodInfo lHandleEventMethod;
        public event EventHandler Event;
        internal string lerror;

        public Bridge(){
            this.lDomain = System.AppDomain.CurrentDomain;
            this.LoadTypes();
            this.lTimeout = 30000;
            lHandleEventMethod = this.GetType().GetMethod("HandleEvent");
            this.timerhotkey = new System.Timers.Timer(100);
            this.timerhotkey.Elapsed += new System.Timers.ElapsedEventHandler(TimerCheckHotKey);
            AppDomain.CurrentDomain.UnhandledException += AppDomain_UnhandledException;
            System.Windows.Forms.Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);

        }
                
        internal void AppDomain_UnhandledException(Object sender, UnhandledExceptionEventArgs e){
            this.lerror = ((Exception)e.ExceptionObject).InnerException.Message;
            this.thread.Abort();
            this.timerhotkey.Stop();
            System.Threading.Thread.CurrentThread.Join();
        }


        internal void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            this.lerror = e.Exception.Message;
            this.thread.Abort();
            this.timerhotkey.Stop();
            System.Threading.Thread.CurrentThread.Join();
        }

        internal void TimerCheckHotKey(object source, System.Timers.ElapsedEventArgs e){
            if ((GetKeyState(0x1b) & 0x8000) != 0) {
                this.timerhotkey.Stop();
                this.thread.Abort();
            }
        }

        internal void HandleEvent(Object sender, EventArgs e){
            if (Event != null){
                Event("", sender, e);
            }
        }

        /// <summary>Waits the provided time (millisecond)</summary>
        /// <param name="pTimeMillisecond"></param>
        public void Wait(int pTimeMillisecond){
            System.Threading.Thread.Sleep(pTimeMillisecond);
        }

        /// <summary>Get/Set the Timeout for asynchronous call</summary>
        public int Timeout{
            get{return this.lTimeout;}
            set{this.lTimeout=value;}
        }

        /// <summary>Get a Type</summary>
        /// <param name="pFullTypeName"></param>
        /// <returns></returns>
        public Type Type(string pFullTypeName){
            System.Type lType;
            lType = this.GetLoadedTypeFromString(pFullTypeName);
            if (lType!=null){
                return new Type(this , lType);
            }else{
                throw new ApplicationException("Type <" + pFullTypeName + "> not found! ");
            }
        }

        /// <summary>Get an assembly by name</summary>
        /// <param name="pAssemblyName"></param>
        /// <returns></returns>
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
                        var ret = e;
                       //System.Console.WriteLine(lTypes[t].FullName);
                    }
                }
            }
        }

        /// <summary>"Load a dll by path</summary>
        /// <param name="pDllPath"></param>
        /// <returns></returns>
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

        /// <summary>Load an assembly from the GAC (C:\WINDOWS\assembly)</summary>
        /// <param name="pAssemblyName"></param>
        /// <param name="pVersion"></param>
        /// <param name="pCulture"></param>
        /// <param name="pPublicKeyToken"></param>
        /// <returns></returns>
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

        /// <summary>Get the list of types</summary>
        /// <returns></returns>
        public System.String[] GetTypesList(){
            string[] lRet = new string[cTypes.Count];
            cTypes.Keys.CopyTo(lRet, 0);
            return lRet;
        }

        /// <summary>Get the list of assemblies</summary>
        /// <returns></returns>
        public System.String[] GetAssembliesList(){
            string[] lRet = new string[cAssemblies.Count];
            cAssemblies.Keys.CopyTo(lRet, 0);
            return lRet;
        }

    }
}
