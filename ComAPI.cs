using Microsoft.Win32;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace TestCOM {

    internal static class ComAPI {
        [DllImport("ole32.DLL")]
        public static extern int CoInitializeSecurity(
            IntPtr securityDescriptor, int cAuth, IntPtr asAuthSvc, IntPtr reserved, uint AuthLevel,
            uint ImpLevel, IntPtr pAuthList, uint Capabilities, IntPtr reserved3);
        [DllImport("ole32.dll")]
        public static extern int CoRegisterClassObject(
            ref Guid rclsid,
            [MarshalAs(UnmanagedType.Interface)] IClassFactory pUnkn,
            int dwClsContext,
            int flags,
            out int lpdwRegister);
        [DllImport("ole32.dll")]
        public static extern int CoRevokeClassObject(int dwRegister);

        public const int RPC_C_AUTHN_LEVEL_PKT_PRIVACY = 6; // Encrypted DCOM communication
        public const int RPC_C_IMP_LEVEL_IDENTIFY = 2;  // No impersonation really required
        public const int CLSCTX_LOCAL_SERVER = 4;
        public const int REGCLS_SINGLEUSE = 0;
        public const int REGCLS_MULTIPLEUSE = 1;
        public const int EOAC_DISABLE_AAA = 0x1000;  // Disable Activate-as-activator
        public const int EOAC_NO_CUSTOM_MARSHAL = 0x2000; // Disable custom marshalling
        public const int EOAC_SECURE_REFS = 0x2;   // Enable secure DCOM references
        public const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
        public const int E_NOINTERFACE = unchecked((int)0x80004002);
        public const string guidIUnknown = "00000000-0000-0000-C000-000000000046";



        public static void RegasmRegisterLocalServer(Type t) {
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

        public static void RegasmUnregisterLocalServer(Type t) {
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

    [ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000001-0000-0000-C000-000000000046")]
    internal interface IClassFactory {
        [PreserveSig]
        int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject);
        [PreserveSig]
        int LockServer(bool fLock);
    }

    internal class ComService<Interface, CoClass> : IClassFactory where CoClass : Interface {
        private Func<object, CoClass> _creator;
        private int _nLockCnt = 0;
        private int _cookie;
        private Semaphore _track = new Semaphore(0, 1);

        class Tracker {
            private ComService<Interface, CoClass> _service;
            public Tracker(ComService<Interface, CoClass> service) {
                service.Lock();
                _service = service;
            }
            ~Tracker() {
                _service.Unlock();
            }
        }

        public ComService(Func<object, CoClass> creator) {
            _creator = creator;
        }

        public int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject) {
            ppvObject = IntPtr.Zero;
            if (pUnkOuter != IntPtr.Zero)
                Marshal.ThrowExceptionForHR(ComAPI.CLASS_E_NOAGGREGATION);
            if (riid == typeof(Interface).GUID || riid == new Guid(ComAPI.guidIUnknown)) {
                ppvObject = Marshal.GetComInterfaceForObject(
                    _creator(new Tracker(this)), typeof(IMyInterface));
            }
            else
                Marshal.ThrowExceptionForHR(ComAPI.E_NOINTERFACE);
            return 0;
        }

        public int LockServer(bool lockIt) {
            if (lockIt)
                Lock();
            else
                Unlock();
            return 0;
        }

        public void Run() {
            int hResult;
#if false
            Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            hResult = ComAPI.CoInitializeSecurity(
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
#endif

            Guid CLSID_MyObject = typeof(CoClass).GUID;
            hResult = ComAPI.CoRegisterClassObject(
                ref CLSID_MyObject,
                this,
                ComAPI.CLSCTX_LOCAL_SERVER,
                ComAPI.REGCLS_SINGLEUSE,
                out _cookie);
            if (hResult != 0)
                Marshal.ThrowExceptionForHR(hResult);
            _nLockCnt = 0;

            try {
                using (
                    var timer = new Timer(
                        new TimerCallback((s) => {
                            GC.Collect();
                        }),
                        null, 5000, 5000
                    )
                ) {
                    _track.WaitOne();
                }
            }
            finally {
                ComAPI.CoRevokeClassObject(_cookie);
            }
        }

        public int Lock() {
            return Interlocked.Increment(ref _nLockCnt);
        }

        public int Unlock() {
            int nRet = Interlocked.Decrement(ref _nLockCnt);
            if (nRet == 0)
                _track.Release();
            return nRet;
        }
    }

}
