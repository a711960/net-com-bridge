// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;
using System.IO;
[assembly: System.Reflection.AssemblyTitle("NetComBridge")]
[assembly: System.Reflection.AssemblyDescription("NetComBridge Type Library")]
[assembly: System.Reflection.AssemblyConfiguration("")]
[assembly: System.Reflection.AssemblyCompany("Florent Breheret")]
[assembly: System.Reflection.AssemblyProduct("NetComBridge")]
[assembly: System.Reflection.AssemblyCopyright("Copyright © Florent Breheret 2012")]
[assembly: System.Reflection.AssemblyTrademark("")]
[assembly: System.Reflection.AssemblyCulture("")]
[assembly: System.Runtime.InteropServices.ComVisible(false)]
[assembly: System.Runtime.InteropServices.Guid("2b7c45c0-8cb3-454c-977a-7a4adc0126cf")]
[assembly: System.Reflection.AssemblyVersion("1.0.8.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.0.8.0")]


namespace NetComBridge
{
    [Guid("FEF31440-3216-497C-92AE-F4961116921A"), ComVisible(true)]
    public interface IAssemblyInfo
    {
        [Description("Get the assembly version")]
        string GetVersion();
        string GetFolder();
    }

    [Description("Class to get informations about the regitered assembly"), ProgId("NetComBridge.AssemblyInfo")]
    [Guid("0C0347D4-6F6D-4ECB-A38F-3FDB1F18FAD2"), ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class AssemblyInfo : IAssemblyInfo
    {
        public string GetVersion()
        {
            try{
                return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }catch (Exception){
                return null;
            }
        }
    
        public string GetFolder()
        {
            try{
                return Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).FullName;
            }catch (Exception){
                return null;
            }
        }

    }
}
