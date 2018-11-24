using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;

namespace ScriptEngine
{
    [ComImport]
    [Guid("00020400-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        int GetTypeInfoCount();

        System.Runtime.InteropServices.ComTypes.ITypeInfo
            GetTypeInfo([MarshalAs(UnmanagedType.U4)] int iTInfo,
                        [MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        int GetIDsOfNames(ref Guid riid,
                          [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames,
                          int cNames,
                          int lcid,
                          [Out][MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

        [PreserveSig]
        int Invoke(int dispIdMember,
                   ref Guid riid,
                   [MarshalAs(UnmanagedType.U4)] int lcid,
                   [MarshalAs(UnmanagedType.U4)] int dwFlags,
                   ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
                   [Out, MarshalAs(UnmanagedType.LPArray)] object[] pVarResult,
                   ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
                   [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
    }


    internal class WdwFlags
    {
        public const int DISPATCH_METHOD = 1;
        public const int DISPATCH_PROPERTYGET = 2;
        public const int DISPATCH_PROPERTYPUT = 4;
        public const int DISPATCH_PROPERTYPUTREF = 8;
    }

    [ComVisible(true)]
    public class DispatchProxy : IDispatch, ICustomQueryInterface, IDisposable
    {
        EngineHost _engineHost = null;

        private Dictionary<string, int> _methodNameIdDictionary = new Dictionary<string, int>();
        private Dictionary<int, string> _invokeNameIdDictionary = new Dictionary<int, string>();

        public object ClassObject { get; private set; }

        public DispatchProxy(EngineHost engineHost, object obj)
        {
            _engineHost = engineHost;
            ClassObject = obj;

            BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Default | BindingFlags.GetProperty | BindingFlags.SetProperty;
            MemberInfo[] members = obj.GetType().GetMembers(bindingFlags);
            int i = 0;
            foreach (MemberInfo mi in members)
            {
                string strUnicode = "";
                if (!ConvertChineseChar.IsChina(mi.Name, ref strUnicode))
                    strUnicode = mi.Name;
                if (_methodNameIdDictionary.ContainsKey(strUnicode))
                    continue;
                _methodNameIdDictionary.Add(strUnicode, i);
                _invokeNameIdDictionary.Add(i, strUnicode);
                i++;
            }
        }

        public int GetTypeInfoCount()
        {
            return 0;
        }

        public ITypeInfo GetTypeInfo([MarshalAs(UnmanagedType.U4)] int iTInfo, [MarshalAs(UnmanagedType.U4)] int lcid)
        {
            return null;
        }

        public int GetIDsOfNames(ref Guid riid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames,
            int cNames,
            int lcid,
            [MarshalAs(UnmanagedType.LPArray), Out] int[] rgDispId)
        {
            System.Console.WriteLine($"IDispatch>>GetIDsOfNames : rgsNames={(rgsNames != null && rgsNames.Length > 0 ? rgsNames[0] : "") } ");

            String strMethodName = rgsNames[0];
            string strUnicode = ConvertChineseChar.UnicodeToChinese(strMethodName);
            object[] arguments = { strUnicode };

            if (!_methodNameIdDictionary.ContainsKey(strUnicode))//不区分大小写匹配
            {
                rgDispId[0] = -1;
            }
            else
            {
                rgDispId[0] = _methodNameIdDictionary[strUnicode];
            }
            return 0;
        }

        private BindingFlags GetInvokeBindFlags(int dwFlags)
        {
            BindingFlags bflags;
            switch (dwFlags)
            {
                case 1: //method
                    bflags = BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static |
                             BindingFlags.Instance | BindingFlags.IgnoreCase | BindingFlags.GetProperty;
                    break;
                case 2: //get property
                    bflags = BindingFlags.Public | BindingFlags.Static | BindingFlags.GetProperty |
                             BindingFlags.GetField | BindingFlags.Instance | BindingFlags.IgnoreCase;
                    break;
                case 4: //set property
                    bflags = BindingFlags.Public | BindingFlags.SetProperty | BindingFlags.Instance |
                             BindingFlags.SetField | BindingFlags.Static | BindingFlags.IgnoreCase;
                    break;
                case 8: //set property
                    bflags = BindingFlags.Public | BindingFlags.Static | BindingFlags.SetProperty |
                             BindingFlags.SetField | BindingFlags.Instance | BindingFlags.IgnoreCase;
                    break;
                default: ////method  
                    bflags = BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static |
                             BindingFlags.Instance | BindingFlags.GetProperty | BindingFlags.GetField | BindingFlags.IgnoreCase;
                    break;
            }
            return bflags;
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            System.Console.WriteLine($"IDispatch>>GetInterface : Guid={iid.ToString()}");

            if (iid.Equals(_engineHost.guid))
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IDispatch), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            ppv = IntPtr.Zero;
            return CustomQueryInterfaceResult.NotHandled;
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        private object[] GetArgumentFromVariant(DISPPARAMS dispParams)
        {
            object[] arguments = null;
            if (dispParams.rgvarg != IntPtr.Zero && dispParams.cArgs != 0)
            {
                arguments = new object[dispParams.cArgs];

                VARENUM vt = VARENUM.VT_BYREF | VARENUM.VT_VARIANT;// 指向Variant对象类型
                int variantSize = 16;// 每个变体地址大小

                // 每个参数的基址
                IntPtr baseIntPtr;
                for (int i = 0; i < dispParams.cArgs; i++)
                {
                    // 第i个参数指针
                    baseIntPtr = IntPtr.Add(dispParams.rgvarg, variantSize * i);

                    // 获取第i个参数的数据类型
                    short vtType = Marshal.ReadInt16(baseIntPtr);
                    if (vtType == (int)vt)
                    {
                        // 如果是变体类型
                        // 偏移8位(2+3*2 )，取得存放指针的内存地址
                        // VARTYPE vt; 占2字节
                        // WORD wReserved1;占2字节
                        // WORD wReserved2;占2字节
                        // WORD wReserved3;占2字节
                        IntPtr intPtr = IntPtr.Add(baseIntPtr, 8);
                        int address = Marshal.ReadInt32(intPtr);
                        arguments[i] = Marshal.GetObjectForNativeVariant((IntPtr)address);
                    }
                    else
                    {
                        arguments[i] = Marshal.GetObjectForNativeVariant(baseIntPtr);
                    }
                }
                Array.Reverse(arguments);
            }

            return arguments;
        }

        public int Invoke(int dispIdMember,
            ref Guid riid,
            [MarshalAs(UnmanagedType.U4)] int lcid,
            [MarshalAs(UnmanagedType.U4)] int dwFlags,
            ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
            [MarshalAs(UnmanagedType.LPArray), Out] object[] pVarResult,
            ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
            [MarshalAs(UnmanagedType.LPArray), Out] IntPtr[] pArgErr)
        {
            object[] arguments = GetArgumentFromVariant(pDispParams);
            if (dwFlags == WdwFlags.DISPATCH_PROPERTYPUT || dwFlags == WdwFlags.DISPATCH_PROPERTYPUTREF)////默认属性Set.Value
            {
                string strArgName;
                _invokeNameIdDictionary.TryGetValue(dispIdMember, out strArgName);
                object obj = TypeUtils.DispatchInvokeMember(strArgName, GetInvokeBindFlags(dwFlags), null, this.ClassObject, arguments);
            }
            else
            {
                if(((dwFlags & WdwFlags.DISPATCH_PROPERTYGET) == WdwFlags.DISPATCH_PROPERTYGET)
                    || ((dwFlags & WdwFlags.DISPATCH_METHOD) == WdwFlags.DISPATCH_METHOD))
                {
                    Console.WriteLine("dwFlags == WdwFlags.DISPATCH_PROPERTYGET");
                }

                string strArgName;
                _invokeNameIdDictionary.TryGetValue(dispIdMember, out strArgName);
                object obj = TypeUtils.DispatchInvokeMember(strArgName, GetInvokeBindFlags(dwFlags), null, this.ClassObject, arguments);
                if (pVarResult != null)
                {
                    pVarResult[0] = obj;
                }
            }
            return 0;
        }
    }
}
