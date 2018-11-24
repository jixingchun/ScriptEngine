using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ScriptEngine
{
    #region 接口定义
    [Guid("BB1A2AE1-A4F9-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScript
    {
        [PreserveSig]
        int SetScriptSite(IActiveScriptSite pass);
        [PreserveSig]
        int GetScriptSite(Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object site);
        [PreserveSig]
        int SetScriptState(ScriptState state);
        [PreserveSig]
        int GetScriptState(out ScriptState scriptState);
        [PreserveSig]
        int Close();
        [PreserveSig]
        int AddNamedItem(string name, ScriptItem flags);
        [PreserveSig]
        int AddTypeLib(Guid typeLib, uint major, uint minor, uint flags);
        [PreserveSig]
        int GetScriptDispatch(string itemName, [MarshalAs(UnmanagedType.IDispatch)] out object dispatch);
        [PreserveSig]
        int GetCurrentScriptThreadID(out uint thread);
        [PreserveSig]
        int GetScriptThreadID(uint win32ThreadId, out uint thread);
        [PreserveSig]
        int GetScriptThreadState(uint thread, out ScriptThreadState state);
        [PreserveSig]
        int InterruptScriptThread(uint thread, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo, uint flags);
        [PreserveSig]
        int Clone(out IActiveScript script);
    }


    [Guid("4954E0D0-FBC7-11D1-8410-006008C3FBFC"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptProperty
    {
        [PreserveSig]
        int GetProperty(int dwProperty, IntPtr pvarIndex, out object pvarValue);
        [PreserveSig]
        int SetProperty(int dwProperty, IntPtr pvarIndex, ref object pvarValue);
    }

    [Guid("BB1A2AE2-A4F9-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptParse32
    {
        [PreserveSig]
        int InitNew();
        [PreserveSig]
        int AddScriptlet(string defaultName, string code, string itemName, string subItemName, string eventName, string delimiter, IntPtr sourceContextCookie, uint startingLineNumber, ScriptText flags, out string name, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        [PreserveSig]
        int ParseScriptText(string code, string itemName, IntPtr context, string delimiter, int sourceContextCookie, uint startingLineNumber, ScriptText flags, out object result, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
    }

    [Guid("C7EF7658-E1EE-480E-97EA-D52CB4D76D17"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptParse64
    {
        [PreserveSig]
        int InitNew();
        [PreserveSig]
        int AddScriptlet(string defaultName, string code, string itemName, string subItemName, string eventName, string delimiter, IntPtr sourceContextCookie, uint startingLineNumber, ScriptText flags, out string name, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        [PreserveSig]
        int ParseScriptText(string code, string itemName, IntPtr context, string delimiter, long sourceContextCookie, uint startingLineNumber, ScriptText flags, out object result, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
    }

    #endregion

    #region 枚举类定义

    [Flags]
    public enum ScriptText
    {
        None = 0,
        DelayExecution = 1,
        IsVisible = 2,
        IsExpression = 32,
        IsPersistent = 64,
        HostManageSource = 128
    }

    [Flags]
    public enum ScriptInfo
    {
        None = 0,
        IUnknown = 1,
        ITypeInfo = 2
    }

    [Flags]
    public enum ScriptItem
    {
        None = 0,
        IsVisible = 2,
        IsSource = 4,
        GlobalMembers = 8,
        IsPersistent = 64,
        CodeOnly = 512,
        NoCode = 1024
    }

    public enum ScriptThreadState
    {
        NotInScript = 0,
        Running = 1
    }

    public enum ScriptState
    {
        Uninitialized = 0,
        Started = 1,
        Connected = 2,
        Disconnected = 3,
        Closed = 4,
        Initialized = 5
    }

    // Variant Type 
    enum VARENUM
    {
        VT_EMPTY = 0,
        VT_NULL = 1,
        VT_I2 = 2,
        VT_I4 = 3,
        VT_R4 = 4,
        VT_R8 = 5,
        VT_CY = 6,
        VT_DATE = 7,
        VT_BSTR = 8,
        VT_DISPATCH = 9,
        VT_ERROR = 10,
        VT_BOOL = 11,
        VT_VARIANT = 12,
        VT_UNKNOWN = 13,
        VT_DECIMAL = 14,
        VT_I1 = 16,
        VT_UI1 = 17,
        VT_UI2 = 18,
        VT_UI4 = 19,
        VT_I8 = 20,
        VT_UI8 = 21,
        VT_INT = 22,
        VT_UINT = 23,
        VT_VOID = 24,
        VT_HRESULT = 25,
        VT_PTR = 26,
        VT_SAFEARRAY = 27,
        VT_CARRAY = 28,
        VT_USERDEFINED = 29,
        VT_LPSTR = 30,
        VT_LPWSTR = 31,
        VT_RECORD = 36,
        VT_INT_PTR = 37,
        VT_UINT_PTR = 38,
        VT_FILETIME = 64,
        VT_BLOB = 65,
        VT_STREAM = 66,
        VT_STORAGE = 67,
        VT_STREAMED_OBJECT = 68,
        VT_STORED_OBJECT = 69,
        VT_BLOB_OBJECT = 70,
        VT_CF = 71,
        VT_CLSID = 72,
        VT_VERSIONED_STREAM = 73,
        VT_BSTR_BLOB = 0xfff,
        VT_VECTOR = 0x1000,
        VT_ARRAY = 0x2000,
        VT_BYREF = 0x4000,
        VT_RESERVED = 0x8000,
        VT_ILLEGAL = 0xffff,
        VT_ILLEGALMASKED = 0xfff,
        VT_TYPEMASK = 0xfff
    };
    #endregion

    #region 错误消息
    [Serializable]
    public class ScriptException : Exception
    {
        public ScriptException()
            : base("Script Exception")
        {
        }

        public ScriptException(string message)
            : base(message)
        {
            this.Description = message;
        }

        public string Description { get; internal set; }
        public int Line { get; internal set; }
        public int Column { get; internal set; }
        public int Number { get; internal set; }
        public string Text { get; internal set; }
    }

    #endregion

    #region 引擎宿主
    public class EngineHost : IDisposable
    {
        #region 常量
        private const int TYPE_E_ELEMENTNOTFOUND = unchecked((int)(0x8002802B));
        private const int E_NOTIMPL = -2147467263;

        /// <summary>
        /// The name of the function used for simple evaluation.
        /// </summary>
        public const string MethodName = "EvalMethod";

        /// <summary>
        /// 默认为VBScript
        /// </summary>
        public const string DefaultLanguage = VBScriptLanguage;

        /// <summary>
        /// VBScript
        /// </summary>
        public const string VBScriptLanguage = "vbscript";

        /// <summary>
        /// 脚本引擎 Engine CLSID
        /// </summary>
        public const string ScriptClsid = "{16d51579-a30b-4c8b-a276-0ff4dc41e755}";

        #endregion

        #region 内部变量
        private IActiveScript _engine;
        private IActiveScriptParse32 _parse32;
        private IActiveScriptParse64 _parse64;
        internal ActiveScriptSite _site;

        public Guid guid = new Guid("00020400-0000-0000-c000-000000000046");
        #endregion

        #region 初始化
        public EngineHost() : this(DefaultLanguage)
        {

        }

        /// <summary> 
        /// 脚本引起初始化
        /// </summary> 
        /// <param name="scriptLanguage">vbscript或者JavaScript</param> 
        public EngineHost(string scriptLanguage)
        {
            if (scriptLanguage == null)
                throw new ArgumentNullException("language");

            Type engineType = Type.GetTypeFromProgID(scriptLanguage, true);
            _engine = Activator.CreateInstance(engineType) as IActiveScript;

            if (_engine == null)
                throw new ArgumentException($"{scriptLanguage} engine is not exist", "scriptLanguage");

            _site = new ActiveScriptSite();
            _engine.SetScriptSite(_site);
            if (IntPtr.Size == 4)
            {
                _parse32 = (IActiveScriptParse32)_engine;
                _parse32.InitNew();
            }
            else
            {
                _parse64 = (IActiveScriptParse64)_engine;
                _parse64.InitNew();
            }
        }
        #endregion

        /// <summary> 
        /// 添加项目名称到脚本引擎
        /// </summary> 
        /// <param name="name">项目名</param> 
        /// <param name="value">项目值，它必须是Com可见</param> 
        public void AddNamedItem(string name, object value)
        {
            if (name == null)
                throw new ArgumentNullException("name");

            DispatchProxy dispatchProxy = new DispatchProxy(this, value);
            _engine.AddNamedItem(name, ScriptItem.IsVisible | ScriptItem.IsSource);

            _site.NamedItems[name] = dispatchProxy;
        }

        #region 编译检查

        /// <summary>
        /// 脚本编译检查
        /// </summary>
        /// <param name="script">脚本内容</param>
        /// <param name="isExpression">是否是表达式</param>
        /// <returns></returns>        
        public List<ScriptException> CompileScript(string script, bool isExpression = false)
        {
            List<ScriptException> scriptExceptions = new List<ScriptException>();

            _engine.SetScriptState(ScriptState.Initialized);

            ScriptText flags = ScriptText.None;
            if (isExpression)
            {
                flags |= ScriptText.IsExpression;
            }

            try
            {
                object result;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
                if (_parse32 != null)
                {
                    _parse32.ParseScriptText(script, null, IntPtr.Zero, null, 0, 0, flags, out result, out exceptionInfo);
                }
                else
                {
                    _parse64.ParseScriptText(script, null, IntPtr.Zero, null, 0, 0, flags, out result, out exceptionInfo);
                }
            }
            catch (Exception ex)
            {
                ScriptException scriptException = new ScriptException();
                scriptException.Description = "ParseScript Error:" + ex;
                scriptExceptions.Add(scriptException);
            }

            if (_site.LastException != null)
            {
                scriptExceptions.Add(_site.LastException);
            }

            return scriptExceptions;
        }
        #endregion


        #region 执行方法

        /// <summary>
        /// 把普通的Class对象包装成Com对象，以便脚本对Class对象进行访问操作
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public DispatchProxy CreateDispatchProxy(object obj)
        {
            DispatchProxy dispatchProxy = new DispatchProxy(this, obj);
            return dispatchProxy;
        }

        /// <summary>
        /// 执行表达式
        /// </summary>
        /// <param name="expression">表达式</param>
        /// <returns></returns>        
        public object EvalExpression(string expression)
        {
            // 执行脚本前，需要使引起处于连接状态
            _engine.SetScriptState(ScriptState.Connected);
            ScriptText flags = ScriptText.None | ScriptText.IsExpression;

            object result;
            try
            {
                System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
                if (_parse32 != null)
                {
                    _parse32.ParseScriptText(expression, null, IntPtr.Zero, null, 0, 0, flags, out result, out exceptionInfo);
                }
                else
                {
                    _parse64.ParseScriptText(expression, null, IntPtr.Zero, null, 0, 0, flags, out result, out exceptionInfo);
                }
            }
            catch
            {
                if (_site.LastException != null)
                    throw _site.LastException;

                throw;
            }

            return result;
        }

        /// <summary>
        /// 执行脚本
        /// </summary>
        /// <param name="script">脚本内容</param>
        /// <param name="methodName">方法名</param>
        /// <param name="arguments">参数</param>
        /// <returns></returns>
        public object EvalMethod(string script,string methodName, params object[] arguments)
        {
            ParsedScript parsedScript = Parse(script);
            return parsedScript.CallMethod(methodName, arguments);
        }

        internal ParsedScript Parse(string text)
        {
            if (text == null)
                throw new ArgumentNullException("text");

            return Parse(text, false);
        }

        private ParsedScript Parse(string script, bool isExpression)
        {
            // 执行脚本前，需要使引起处于连接状态
            _engine.SetScriptState(ScriptState.Connected);

            ScriptText flags = ScriptText.None;
            if (isExpression)
            {
                flags |= ScriptText.IsExpression;
            }

            try
            {
                object result;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
                if (_parse32 != null)
                {
                    _parse32.ParseScriptText(script, null, IntPtr.Zero, null, 0, 0, flags, out result, out exceptionInfo);
                }
                else
                {
                    _parse64.ParseScriptText(script, null, IntPtr.Zero, null, 0, 0, flags, out result, out exceptionInfo);
                }
            }
            catch
            {
                if (_site.LastException != null)
                    throw _site.LastException;

                throw;
            }

            object dispatch;
            _engine.GetScriptDispatch(null, out dispatch);
            ParsedScript parsed = new ParsedScript(this, dispatch);
            return parsed;
        }
        #endregion

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            if (_parse32 != null)
            {
                Marshal.ReleaseComObject(_parse32);
                _parse32 = null;
            }

            if (_parse64 != null)
            {
                Marshal.ReleaseComObject(_parse64);
                _parse64 = null;
            }

            if (_engine != null)
            {
                Marshal.ReleaseComObject(_engine);
                _engine = null;
            }
        }
    }
    
    #endregion



}
