using System;
using System.ComponentModel;
using System.Diagnostics;
using System.ServiceProcess;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Reflection;

namespace TestCOM {
    //
    // .NET class, interface exposed through DCOM
    //

    // exposed COM interface
    [Guid(MyService.guidIMyInterface), ComVisible(true)]
    public interface IMyInterface {
        string GetDateTime(string prefix);
    }

    // exposed COM class
    [Guid(MyService.guidMyClass), ComVisible(true)]
    public class CMyClass : IMyInterface {
        // Print date & time and the current EXE name
        public string GetDateTime(string prefix) {
            Process currentProcess = Process.GetCurrentProcess();
            return string.Format("{0}: { 1} [server-side COM call executed on {2}]",
                prefix, DateTime.Now, currentProcess.MainModule.ModuleName);
        }

        [ComRegisterFunction(), EditorBrowsable(EditorBrowsableState.Never)]
        void Register(Type t) {
            try {
                if (t == null)
                    throw new ArgumentNullException("The CLR type must be specified.", "t");
                RegistryKey keyCLSID = Registry.ClassesRoot.OpenSubKey("CLSID\\" + t.GUID.ToString("B"), /*writable*/ true);
                try {
                    // Remove the auto-generated InprocServer32 key after registration
                    // (REGASM puts it there but we are going out-of-proc).
                    keyCLSID.DeleteSubKeyTree("InprocServer32");

                    // Create "LocalServer32" under the CLSID key
                    RegistryKey subkey = keyCLSID.CreateSubKey("LocalServer32");
                    try {
                        subkey.SetValue("", Assembly.GetExecutingAssembly().Location, RegistryValueKind.String);
                    }
                    finally {
                        subkey.Dispose();
                    }
                }
                finally {
                    keyCLSID.Dispose();
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        [ComUnregisterFunction(), EditorBrowsable(EditorBrowsableState.Never)]
        static void Unregister(Type t) {
            try {
                if (t == null)
                    throw new ArgumentNullException("The CLR type must be specified.", "t");
                Registry.ClassesRoot.DeleteSubKeyTree("CLSID\\" + t.GUID.ToString("B"));
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
    }

    //
    // My hosting Windows service
    //
    internal class MyService : ServiceBase {
        public MyService() {
            // Initialize COM security
            Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
             uint hResult = ComAPI.CoInitializeSecurity(
                IntPtr.Zero, // Add here your Security descriptor
                -1,
                IntPtr.Zero,
                IntPtr.Zero,
                ComAPI.RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
                ComAPI.RPC_C_IMP_LEVEL_IDENTIFY,
                IntPtr.Zero,
                ComAPI.EOAC_DISABLE_AAA | ComAPI.EOAC_SECURE_REFS | ComAPI.EOAC_NO_CUSTOM_MARSHAL,
                IntPtr.Zero);
            if (hResult != 0)
                throw new ApplicationException("CoIntializeSecurity failed" + hResult.ToString("X"));
        }

        // The main entry point for the process
        static void Main() {
            Run(new ServiceBase[] { new MyService() });
        }

        /// 
        /// On start, register the COM class factory
        /// 
        protected override void OnStart(string[] args) {
            Guid CLSID_MyObject = new Guid(guidMyClass);
            uint hResult = ComAPI.CoRegisterClassObject(
                ref CLSID_MyObject,
                new MyClassFactory(),
                ComAPI.CLSCTX_LOCAL_SERVER,
                ComAPI.REGCLS_MULTIPLEUSE,
                out _cookie);
            if (hResult != 0)
                throw new ApplicationException("CoRegisterClassObject failed" + hResult.ToString("X"));
        }

        /// 
        /// On stop, remove the COM class factory registration
        /// 
        protected override void OnStop() {
            if (_cookie != 0)
                ComAPI.CoRevokeClassObject(_cookie);
        }

        private int _cookie = 0;

        //
        // Public constants
        //
        public const string serviceName = "MyService";
        public const string guidIMyInterface = "e88d15a5-0510-4115-9aee-a8421c96decb";
        public const string guidMyClass = "f681abd0-41de-46c8-9ed3-d0f4eba19891";
    }

    //
    // Standard installer
    //
    [RunInstaller(true)]
    public class MyServiceInstaller : System.Configuration.Install.Installer {
        public MyServiceInstaller() {
            processInstaller = new ServiceProcessInstaller();
            serviceInstaller = new ServiceInstaller();
            // Add a new service running under Local SYSTEM
            processInstaller.Account = ServiceAccount.LocalSystem;
            serviceInstaller.StartType = ServiceStartMode.Manual;
            serviceInstaller.ServiceName = MyService.serviceName;
            Installers.Add(serviceInstaller);
            Installers.Add(processInstaller);
        }
        private ServiceInstaller serviceInstaller;
        private ServiceProcessInstaller processInstaller;
    }

    //
    // Internal COM Stuff
    //

    /// 
    /// P/Invoke calls
    /// 

    internal class ComAPI {
        [DllImport("ole32.DLL")]
        public static extern uint CoInitializeSecurity(
            IntPtr securityDescriptor, int cAuth, IntPtr asAuthSvc, IntPtr reserved, uint AuthLevel,
            uint ImpLevel, IntPtr pAuthList, uint Capabilities, IntPtr reserved3);
        [DllImport("ole32.dll")]
        public static extern uint CoRegisterClassObject(
            ref Guid rclsid,
            [MarshalAs(UnmanagedType.Interface)] IClassFactory pUnkn,
            int dwClsContext,
            int flags,
            out int lpdwRegister);
        [DllImport("ole32.dll")]
        public static extern uint CoRevokeClassObject(int dwRegister);

        public const int RPC_C_AUTHN_LEVEL_PKT_PRIVACY = 6; // Encrypted DCOM communication
        public const int RPC_C_IMP_LEVEL_IDENTIFY = 2;  // No impersonation really required
        public const int CLSCTX_LOCAL_SERVER = 4;
        public const int REGCLS_MULTIPLEUSE = 1;
        public const int EOAC_DISABLE_AAA = 0x1000;  // Disable Activate-as-activator
        public const int EOAC_NO_CUSTOM_MARSHAL = 0x2000; // Disable custom marshalling
        public const int EOAC_SECURE_REFS = 0x2;   // Enable secure DCOM references
        public const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
        public const int E_NOINTERFACE = unchecked((int)0x80004002);
        public const string guidIClassFactory = "00000001-0000-0000-C000-000000000046";
        public const string guidIUnknown = "00000000-0000-0000-C000-000000000046";
    }

    /// 
    /// IClassFactory declaration
    /// 

    [ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid(ComAPI.guidIClassFactory)]
    internal interface IClassFactory {
        [PreserveSig]
        int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject);
        [PreserveSig]
        int LockServer(bool fLock);
    }

    /// 
    /// My Class factory implementation
    /// 
    internal class MyClassFactory : IClassFactory {
        public int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject) {
            ppvObject = IntPtr.Zero;
            if (pUnkOuter != IntPtr.Zero)
                Marshal.ThrowExceptionForHR(ComAPI.CLASS_E_NOAGGREGATION);
            if (riid == new Guid(MyService.guidIMyInterface) || riid == new Guid(ComAPI.guidIUnknown)) {
                //
                // Create the instance of my .NET object
                //
                ppvObject = Marshal.GetComInterfaceForObject(
                    new CMyClass(), typeof(IMyInterface));
            }
            else
                Marshal.ThrowExceptionForHR(ComAPI.E_NOINTERFACE);
            return 0;
        }
        public int LockServer(bool lockIt) {
            return 0;
        }
    }

}
