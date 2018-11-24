using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ScriptEngine
{
    internal sealed class ParsedScript : IDisposable
    {
        private object _dispatch;
        private readonly EngineHost _engine;

        internal ParsedScript(EngineHost engine, object dispatch)
        {
            _engine = engine;
            _dispatch = dispatch;
        }

        internal object CallMethod(string methodName, params object[] arguments)
        {
            if (_dispatch == null)
                throw new InvalidOperationException();

            if (methodName == null)
                throw new ArgumentNullException("methodName");

            try
            {
                return _dispatch.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, _dispatch, arguments);
            }
            catch
            {
                if (_engine._site.LastException != null)
                    throw _engine._site.LastException;

                throw;
            }
        }

        void IDisposable.Dispose()
        {
            if (_dispatch != null)
            {
                Marshal.ReleaseComObject(_dispatch);
                _dispatch = null;
            }
        }
    }
}
