using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace ScriptEngine
{
    #region 接口定义
    [Guid("DB01A1E3-A42B-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptSite
    {
        [PreserveSig]
        int GetLCID(out int lcid);
        [PreserveSig]
        int GetItemInfo(string name, ScriptInfo returnMask, [Out, MarshalAs(UnmanagedType.IUnknown)] out object item, IntPtr typeInfo);
        [PreserveSig]
        int GetDocVersionString(out string version);
        [PreserveSig]
        int OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        [PreserveSig]
        int OnStateChange(ScriptState scriptState);
        [PreserveSig]
        int OnScriptError(IActiveScriptError scriptError);
        [PreserveSig]
        int OnEnterScript();
        [PreserveSig]
        int OnLeaveScript();
    }

    [Guid("EAE1BA61-A4ED-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptError
    {
        [PreserveSig]
        int GetExceptionInfo(out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        [PreserveSig]
        int GetSourcePosition(out uint sourceContext, out int lineNumber, out int characterPosition);
        [PreserveSig]
        int GetSourceLineText(out string sourceLine);
    }
    #endregion

    internal class ActiveScriptSite : IActiveScriptSite
    {
        #region 常量
        private const int TYPE_E_ELEMENTNOTFOUND = unchecked((int)(0x8002802B));
        private const int E_NOTIMPL = -2147467263;

        #endregion

        /// <summary>
        /// 保留最后一次错误信息
        /// </summary>
        internal ScriptException LastException;

        /// <summary>
        /// 项目名、值信息
        /// </summary>
        internal Dictionary<string, object> NamedItems = new Dictionary<string, object>();

        int IActiveScriptSite.GetItemInfo(string name, ScriptInfo returnMask,
            [Out, MarshalAs(UnmanagedType.IUnknown)] out object item, IntPtr typeInfo)
        {
            item = IntPtr.Zero;
            if ((returnMask & ScriptInfo.ITypeInfo) == ScriptInfo.ITypeInfo)
                return E_NOTIMPL;

            object value;
            if (!NamedItems.TryGetValue(name, out value))
                return TYPE_E_ELEMENTNOTFOUND;

            item = value;
            return 0;
        }


        int IActiveScriptSite.OnScriptError(IActiveScriptError scriptError)
        {
            string sourceLine = null;
            try
            {
                scriptError.GetSourceLineText(out sourceLine);
            }
            catch
            {
                // 执行GetSourceLineText有时会发生错误，所以做特殊处理
            }
            uint sourceContext;
            int lineNumber;
            int characterPosition;
            scriptError.GetSourcePosition(out sourceContext, out lineNumber, out characterPosition);
            lineNumber++;
            characterPosition++;
            System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
            scriptError.GetExceptionInfo(out exceptionInfo);

            string message;
            if (!string.IsNullOrEmpty(sourceLine))
            {
                message = "Script exception: {1}. Error number {0} (0x{0:X8}): {2} at line {3}, column {4}. Source line: '{5}'.";
            }
            else
            {
                message = "Script exception: {1}. Error number {0} (0x{0:X8}): {2} at line {3}, column {4}.";
            }
            LastException = new ScriptException(string.Format(message, exceptionInfo.scode, exceptionInfo.bstrSource, exceptionInfo.bstrDescription, lineNumber, characterPosition, sourceLine));
            LastException.Column = characterPosition;
            LastException.Description = exceptionInfo.bstrDescription;
            LastException.Line = lineNumber;
            LastException.Number = exceptionInfo.scode;
            LastException.Text = sourceLine;
            return 0;
        }

        int IActiveScriptSite.OnEnterScript()
        {
            LastException = null;
            return 0;
        }

        int IActiveScriptSite.OnLeaveScript()
        {
            return 0;
        }

        int IActiveScriptSite.GetLCID(out int lcid)
        {
            lcid = Thread.CurrentThread.CurrentCulture.LCID;
            return 0;
        }

        int IActiveScriptSite.GetDocVersionString(out string version)
        {
            version = "1.0";
            return 0;
        }

        int IActiveScriptSite.OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo)
        {
            return 0;
        }

        int IActiveScriptSite.OnStateChange(ScriptState scriptState)
        {
            return 0;
        }
    }
}
