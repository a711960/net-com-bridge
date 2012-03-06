using System;
using System.Runtime.InteropServices;

namespace NetComBridgeLib
{
    /// <summary>
    /// Class reffering to methods
    /// </summary>
    [Guid("68db4aaa-ff84-46f9-84f2-b4dc21a00a44")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Method : IMethod
    {

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern short GetKeyState(int virtualKeyCode);

        private NetComBridge lBridge;
        private System.Type lType;
        private System.String lMethodName;
        private System.Object lInstance;
        private object[] lArguments;
        private System.Reflection.MethodInfo lMethod;
        private Instance lReturnInstance;
        System.Threading.Thread thread;
        System.Timers.Timer timerhotkey;


        internal Method(NetComBridge netComBridge, System.Object pInstance, System.Type pType, System.String pMethodName){
            this.lBridge = netComBridge;
            this.lInstance = pInstance;
            this.lType = pType;
            this.lMethodName = pMethodName;
            this.timerhotkey = new System.Timers.Timer(100);
            this.timerhotkey.Elapsed += new System.Timers.ElapsedEventHandler(TimerCheckHotKey);
        }

        private void TimerCheckHotKey(object source, System.Timers.ElapsedEventArgs e){
            if ((GetKeyState(0x1b) & 0x8000) != 0) {
                this.timerhotkey.Stop();
                this.thread.Abort();
            }
        }

        public Instance Invoke0(){
            return this.InvokeSynchrone();
        }

        public Instance Invoke1(object pArgument1){
            return this.InvokeSynchrone(pArgument1);
        }

        public Instance Invoke2(object pArgument1, object pArgument2){
            return this.InvokeSynchrone(pArgument1, pArgument2);
        }

        public Instance Invoke3(object pArgument1, object pArgument2, object pArgument3){
            return this.InvokeSynchrone(pArgument1, pArgument2, pArgument3);
        }

        public Instance Invoke4(object pArgument1, object pArgument2, object pArgument3, object pArgument4){
            return this.InvokeSynchrone(pArgument1, pArgument2, pArgument3, pArgument4);
        }

        public Instance Invoke5(object pArgument1, object pArgument2, object pArgument3, object pArgument4, object pArgument5){
            return this.InvokeSynchrone(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5);
        }


        public Instance InvokeM(ref object[] pArguments){
            return this.InvokeMethod(ref pArguments, true);
        }

        public Instance InvokeSynchrone(params object[] pArguments){
            return this.InvokeMethod(ref pArguments, true);
        }

        public Instance InvokeAsynchrone(ref object[] pArguments){
            return this.InvokeMethod(ref pArguments, false);
        }

        public Instance InvokeMethod(ref object[] pArguments, bool pSynchrone){
            //Convert Generic Instance types to real types
            System.Type[] lTypes = new System.Type[pArguments.Length];
            lArguments = new object[pArguments.Length];
            for (int i = 0; i < pArguments.Length; i++){
                if (pArguments[i] is Instance){
                    lTypes[i] = ((Instance)pArguments[i]).BaseType;
                    lArguments[i] = ((Instance)pArguments[i]).Value;
                }else{
                    lTypes[i] = pArguments[i].GetType();
                    lArguments[i] = pArguments[i];
                }
            }
            //Find method 
            lMethod = this.lType.GetMethod(this.lMethodName, lTypes);
            if (lMethod == null){
                System.Type[] lInterfaces = this.lType.GetInterfaces();
                foreach(System.Type lInterface in lInterfaces ){
                    lMethod = lInterface.GetMethod(this.lMethodName, lTypes);
                    if (lMethod != null) break;
                }
            }
            if (lMethod == null) throw new ApplicationException("Method <" + this.lMethodName + "> not found! ");
            if(lMethod.IsStatic==false && lInstance==null) throw new ApplicationException("Can't invoke method <" + this.lMethodName + ">. Type <" + lType.FullName + "> is not instantiated ");
            //Invoke method
            this.lReturnInstance = new Instance(this.lBridge);
            this.thread = new System.Threading.Thread(new System.Threading.ThreadStart(InvokeProc));
            this.thread.Start();

            this.timerhotkey.Start();
            if (pSynchrone){
                bool succed = this.thread.Join(this.lBridge.Timeout);
                this.timerhotkey.Stop();
                if (!succed) throw new ApplicationException("Timeout reached while invoking method <" + this.lMethodName + "> !   ");
                if (this.lReturnInstance.lErrorMessage != null) throw new ApplicationException(this.lReturnInstance.lErrorMessage);
            }
            return this.lReturnInstance;      
        }

        private void InvokeProc(){
            this.lReturnInstance.lIsReady = false;
            System.Object lRet=null;
            try{
                lRet= this.lMethod.Invoke(this.lInstance, this.lArguments);
                if(lRet!=null){
                    this.lReturnInstance.lInstance=lRet;
                    this.lReturnInstance.lType = lRet.GetType();
                }
            }catch(ApplicationException e){
                this.lReturnInstance.lErrorMessage = "Method <" + this.lMethodName + "> invocation failed ! \r\n" + e.InnerException.Message;
            }
            this.timerhotkey.Stop();
            this.lReturnInstance.lIsReady = true;
        }

    }
}
