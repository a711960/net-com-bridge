using System;
using System.Runtime.InteropServices;

namespace NetComBridge
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

        private Bridge lBridge;
        private System.Type lType;
        private System.String lMethodName;
        private System.Object lInstance;
        private object[] lArguments;
        private System.Reflection.MethodInfo lMethod;
        private Instance lReturnInstance;
        System.Threading.Thread thread;
        System.Timers.Timer timerhotkey;


        internal Method(Bridge pBridge, System.Object pInstance, System.Type pType, System.String pMethodName){
            this.lBridge = pBridge;
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

        /// <summary>Invoke the method and wait the end of execution</summary>
        /// <param name="pArgument1"></param>
        /// <param name="pArgument2"></param>
        /// <param name="pArgument3"></param>
        /// <param name="pArgument4"></param>
        /// <param name="pArgument5"></param>
        /// <param name="pArgument6"></param>
        /// <returns></returns>
        public Instance Invok([Optional][DefaultParameterValue(null)]object pArgument1, [Optional][DefaultParameterValue(null)]object pArgument2, 
            [Optional][DefaultParameterValue(null)]object pArgument3, [Optional][DefaultParameterValue(null)]object pArgument4,
            [Optional][DefaultParameterValue(null)]object pArgument5, [Optional][DefaultParameterValue(null)]object pArgument6){

            if(object.Equals(pArgument1,null)){
                return this.InvokeSynchroneParams();
            }else if (object.Equals(pArgument2, null)){
                return this.InvokeSynchroneParams(pArgument1);
            }else if (object.Equals(pArgument3, null)){
                return this.InvokeSynchroneParams(pArgument1, pArgument2);
            }else if (object.Equals(pArgument4, null)){
                return this.InvokeSynchroneParams(pArgument1, pArgument2, pArgument3);
            }else if (object.Equals(pArgument5, null)){
                return this.InvokeSynchroneParams(pArgument1, pArgument2, pArgument3, pArgument4);
            }else if (object.Equals(pArgument6, null)){
                return this.InvokeSynchroneParams(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5);
            }else{
                return this.InvokeSynchroneParams(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5, pArgument6);
            }
        }

        /// <summary>Invoke the method but don't wait the end of execution</summary>
        /// <param name="pArgument1"></param>
        /// <param name="pArgument2"></param>
        /// <param name="pArgument3"></param>
        /// <param name="pArgument4"></param>
        /// <param name="pArgument5"></param>
        /// <param name="pArgument6"></param>
        /// <returns></returns>
        public Instance InvokAsync([Optional][DefaultParameterValue(null)]object pArgument1, [Optional][DefaultParameterValue(null)]object pArgument2, 
            [Optional][DefaultParameterValue(null)]object pArgument3, [Optional][DefaultParameterValue(null)]object pArgument4,
            [Optional][DefaultParameterValue(null)]object pArgument5, [Optional][DefaultParameterValue(null)]object pArgument6){

            if(object.Equals(pArgument1,null)){
                return this.InvokeAsynchroneParams();
            }else if (object.Equals(pArgument2, null)){
                return this.InvokeAsynchroneParams(pArgument1);
            }else if (object.Equals(pArgument3, null)){
                return this.InvokeAsynchroneParams(pArgument1, pArgument2);
            }else if (object.Equals(pArgument4, null)){
                return this.InvokeAsynchroneParams(pArgument1, pArgument2, pArgument3);
            }else if (object.Equals(pArgument5, null)){
                return this.InvokeAsynchroneParams(pArgument1, pArgument2, pArgument3, pArgument4);
            }else if (object.Equals(pArgument6, null)){
                return this.InvokeAsynchroneParams(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5);
            }else{
                return this.InvokeAsynchroneParams(pArgument1, pArgument2, pArgument3, pArgument4, pArgument5, pArgument6);
            }
        }

        /// <summary>Same as "Invok" but with an arguments array as parameter</summary>
        /// <param name="pArguments"></param>
        /// <returns></returns>
        public Instance InvokT(ref object[] pArguments){
            return this.InvokeMethod(ref pArguments, true);
        }

        /// <summary>Same as "InvokAsync" but with an arguments array as parameter</summary>
        /// <param name="pArguments"></param>
        /// <returns></returns>
        public Instance InvokAsyncT(ref object[] pArguments){
            return this.InvokeMethod(ref pArguments, false);
        }

        private Instance InvokeSynchroneParams(params object[] pArguments){
            return this.InvokeMethod(ref pArguments, true);
        }

        private Instance InvokeAsynchroneParams(params object[] pArguments){
            return this.InvokeMethod(ref pArguments, false);
        }

        private Instance InvokeMethod(ref object[] pArguments, bool pSynchrone){
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
