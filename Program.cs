using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace TestCOM {

    [Guid("F919122A-CB04-45C7-89F6-E661C31018D2"), ComVisible(true)]
    public interface IOtherObject {
        string Test { get; }
    }
    [Guid("1563CA44-DD72-4261-91CC-8D6C71F537BB"), ComVisible(true)]
    public class OtherObject : IOtherObject {
        public string Test { get { return "TEXT"; } }
        internal object token;
    }


    [Guid("e88d15a5-0510-4115-9aee-a8421c96decb"), ComVisible(true)]
    public interface IMyInterface {
        string GetDateTime(string prefix);
        IOtherObject GetObj();
    }

    [Guid("f681abd0-41de-46c8-9ed3-d0f4eba19891"), ComVisible(true)]
    public class CMyClass : IMyInterface {
        public string GetDateTime(string prefix) {
            Process currentProcess = Process.GetCurrentProcess();
            return string.Format("{0}: {1} [server-side COM call executed on {2}]",
                prefix, DateTime.Now, currentProcess.MainModule.ModuleName);
        }
        public IOtherObject GetObj() {
            return new OtherObject() { token = token };
        }

        [ComRegisterFunction] static void Register(Type t) { ComAPI.RegasmRegisterLocalServer(t); }
        [ComUnregisterFunction] static void Unregister(Type t) { ComAPI.RegasmUnregisterLocalServer(t); }
        internal object token;
    }

    
    internal class MyService {
        static void Main() {
            new ComService<IMyInterface, CMyClass>((t) => new CMyClass { token = t }).Run();
        }
    }

}
