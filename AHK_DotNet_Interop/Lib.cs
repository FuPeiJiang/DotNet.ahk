using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Collections;
using System.Reflection;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Loader;
using System.IO.Pipes;
using System.Text;

namespace AHK_DotNet_Interop
{
    public static class Lib
    {
        public static bool DebugWriteLineEnabled { get; set; } = true;

        public static void DebugWriteLine(string s)
        {
            if (DebugWriteLineEnabled)
            {
                Console.WriteLine(s);
            }
        }

        public static int Hello(IntPtr arg, int argLength)
        {
            Console.WriteLine("hello world");
            var ptr = Marshal.GetComInterfaceForObject(new Wrapper(_list), typeof(IDispatch));
            Marshal.WriteIntPtr(arg, ptr);
            return 0;
        }

        public static int GetClass(string FullName, IntPtr out_IDispatch)
        {
            try
            {
                Type? found_type = AppDomain.CurrentDomain.GetAssemblies().SelectMany(x => x.GetTypes()).FirstOrDefault(t => t.FullName == FullName);
                if (found_type == null)
                {
                    return 1;
                }
                Marshal.WriteIntPtr(out_IDispatch, Marshal.GetComInterfaceForObject(new Wrapper(found_type), typeof(IDispatch)));
                return 0;
            }
            catch (Exception e)
            {
                DebugWriteLine(e.ToString());
                throw;
            }
        }

        public static int LoadAssembly([MarshalAs(UnmanagedType.LPWStr)] string path, IntPtr out_IDispatch)
        {
            try
            {
                Assembly assembly = AssemblyLoadContext.Default.LoadFromAssemblyPath(path);
                Marshal.WriteIntPtr(out_IDispatch, Marshal.GetComInterfaceForObject(new Wrapper(assembly), typeof(IDispatch)));
                return 0;
            }
            catch (Exception e)
            {
                DebugWriteLine(e.ToString());
                throw;
            }
        }

        public static int CompileAssembly([MarshalAs(UnmanagedType.LPWStr)] string code, [MarshalAs(UnmanagedType.LPWStr)] string assemblyName, [MarshalAs(UnmanagedType.LPWStr)] string externalReferences, IntPtr out_IDispatch)
        {
            try
            {
                using var pipe = new NamedPipeClientStream(
                    ".", // local machine
                    "CSharp-Compiler-Wrapper",
                    PipeDirection.InOut,
                    PipeOptions.None
                );

                pipe.Connect();

                using var reader = new BinaryReader(pipe, Encoding.UTF8, leaveOpen: true);
                using var writer = new BinaryWriter(pipe, Encoding.UTF8, leaveOpen: true);

                writer.Write(assemblyName);
                writer.Write(externalReferences);
                writer.Write(code);

                string success = reader.ReadString();
                switch (success)
                {
                    case "success":
                        {
                            long length = reader.Read7BitEncodedInt64();
                            byte[] payload = reader.ReadBytes((int)length); // or manual loop if large

                            using var ms = new MemoryStream(payload);
                            Assembly assembly = AssemblyLoadContext.Default.LoadFromStream(ms);
                            Marshal.WriteIntPtr(out_IDispatch, Marshal.GetComInterfaceForObject(new Wrapper(assembly), typeof(IDispatch)));
                            return 0;
                        }
                    case "failure":
                        DebugWriteLine(reader.ReadString());
                        break;
                }
                Marshal.WriteIntPtr(out_IDispatch, 0);
                return 1;
            }
            catch (Exception e)
            {
                DebugWriteLine(e.ToString());
                throw;
            }
        }

        public static Array CreateArray(Type elementType, params object?[] values)
        {
            Array arr = Array.CreateInstance(elementType, values.Length);

            if (elementType == typeof(bool))
            {
                for (int i = 0; i < values.Length; i++)
                {
                    object? val = values[i];
                    if (val != null && val.GetType() == typeof(int) && ((int)val == 0 || (int)val == 1))
                    {
                        arr.SetValue((int)val == 1, i);
                    }
                    else
                    {
                        throw new ArgumentException($"Invalid value for bool array, got: {val}");
                    }
                }
            }
            else
            {
                for (int i = 0; i < values.Length; i++)
                {
                    arr.SetValue(values[i], i); // throws if incompatible
                }
            }

            return arr;
        }

        static ArrayList _list = ["1", "2", "3", "4"];

        public static void Main()
        {
            IntPtr pointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(IntPtr)));

            Hello(pointer, 0);
            // Get all assemblies loaded into the current application domain
            Assembly[] loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();

            Console.WriteLine("List of assemblies loaded in current appdomain:");

            // Iterate over the assemblies and display their names
            foreach (Assembly assembly in loadedAssemblies)
            {
                Console.WriteLine(assembly.GetName().FullName);
            }
        }
    }

    [Guid("00020400-0000-0000-C000-000000000046")]
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatch
    {
        [PreserveSig]
        uint GetTypeInfoCount(out uint pctinfo);

        [PreserveSig]
        uint GetTypeInfo(uint iTInfo, int lcid, out IntPtr info);

        [PreserveSig]
        uint GetIDsOfNames(
            ref Guid iid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)]
            string[] names,
            uint cNames,
            int lcid,
            [Out]
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4, SizeParamIndex = 2)]
            int[] rgDispId);

        [PreserveSig]
        uint Invoke(
            int dispIdMember,
            ref Guid riid,
            int lcid,
            InvokeFlags wFlags,
            DISPPARAMS pDispParams,
            IntPtr VarResult,
            IntPtr pExcepInfo,
            IntPtr puArgErr);
    }

    [Flags]
    public enum InvokeFlags : short
    {
        DISPATCH_METHOD = 1,
        DISPATCH_PROPERTYGET = 2,
        DISPATCH_PROPERTYPUT = 4,
        DISPATCH_PROPERTYPUTREF = 8
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("B196B283-BAB4-101A-B69C-00AA00341D07")]
    internal interface IProvideClassInfo
    {
        [PreserveSig]
        uint GetClassInfo(out IntPtr info);
    }

    public class Wrapper : IDispatch, IProvideClassInfo, ITypeInfo
    {
        private const int DISPID_UNKNOWN = -1;
        private const int DISPID_NEWENUM = -4;
        private const uint DISP_E_UNKNOWNNAME = 0x80020006;
        private const uint DISP_E_MEMBERNOTFOUND = 0x80020003;

        private object _obj;
        private Type _type;

        private Dictionary<Type, TypeEntry> _type_entries = [];

        List<List<object>?> _methods_list;
        Dictionary<string, ListAndIdx> _methods_dict;


        struct TypeEntry
        {
            public List<List<object>?> methods_list;
            public Dictionary<string, ListAndIdx> methods_dict;
        }

        struct ListAndIdx
        {
            public List<object> list; // this needs not be stored in _methods_dict
            public int idx;
        }

        public Wrapper(object obj) : this(obj, obj.GetType()) { }
        public Wrapper(Type type) : this(type, type) { }

        public Wrapper(object obj, Type type)
        {
            _obj = obj;
            _type = type;

            ref var type_entry_ref = ref CollectionsMarshal.GetValueRefOrAddDefault(_type_entries, type, out bool type_entry_exists);
            if (type_entry_exists)
            {
                _methods_list = type_entry_ref.methods_list;
                _methods_dict = type_entry_ref.methods_dict;
            }
            else
            {
                _methods_list = [];
                _methods_dict = new(StringComparer.OrdinalIgnoreCase);
                type_entry_ref = new TypeEntry
                {
                    methods_list = _methods_list,
                    methods_dict = _methods_dict
                };
                HashSet<MethodInfo> properties = new();
                foreach (var item in type.GetProperties())
                {
                    MethodInfo? method = item.GetMethod;
                    if (method != null)
                    {
                        properties.Add(method);
                    }
                    method = item.SetMethod;
                    if (method != null)
                    {
                        properties.Add(method);
                    }
                }
                int idx_counter = 1;
                // Console.WriteLine("++++++++++++++++++++");
                // Console.WriteLine(_type.ToString());
                // foreach (var method in type.GetMethods(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
                foreach (var method in type.GetMethods())
                {
                    // if (method.IsStatic)
                    // {
                    //     continue;
                    // }
                    var name = method.Name;
                    if (properties.Contains(method))
                    {
                        name = name[4..];
                    }

                    // var parameters = method.GetParameters();
                    // var parameterDescriptions = string.Join
                    //     (", ", method.GetParameters()
                    //                 .Select(x => x.ParameterType + " " + x.Name)
                    //                 .ToArray());
                    // Console.WriteLine("{0} {1} ({2}) {3}",
                    //                 method.ReturnType,
                    //                 method.Name,
                    //                 parameterDescriptions,
                    //                 method.IsSpecialName ? "special" : "");

                    ref var itemRef = ref CollectionsMarshal.GetValueRefOrAddDefault(_methods_dict, name, out bool exists);
                    if (!exists)
                    {
                        itemRef = new ListAndIdx { list = [], idx = idx_counter };
                        ++idx_counter;
                    }
                    itemRef.list.Add(method);
                }
                if (_methods_dict.TryGetValue("Item", out var value))
                {
                    _methods_list.Add(value.list);
                }
                else if (_obj is Type)
                {
                    _methods_list.Add(_type.GetConstructors().Cast<object>().ToList());
                }
                else
                {
                    _methods_list.Add(null);
                }
                foreach (var item in _methods_dict.Values)
                {
                    _methods_list.Add(item.list);
                }

                foreach (var field in type.GetFields())
                {
                    ref var itemRef = ref CollectionsMarshal.GetValueRefOrAddDefault(_methods_dict, field.Name, out bool exists);
                    if (!exists)
                    {
                        itemRef = new ListAndIdx { list = [], idx = idx_counter };
                        ++idx_counter;
                    }
                    itemRef.list.Add(field);
                    _methods_list.Add(itemRef.list);
                }
            }
        }

        public uint GetTypeInfoCount(out uint pctinfo)
        {
            Lib.DebugWriteLine("GetTypeInfoCount");
            pctinfo = 0;
            return 0;
        }

        public uint GetTypeInfo(uint iTInfo, int lcid, out nint info)
        {
            Lib.DebugWriteLine("GetTypeInfo");
            info = 0;
            return 0;
        }

        public uint GetIDsOfNames(ref Guid iid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)] string[] names, uint cNames, int lcid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4, SizeParamIndex = 2), Out] int[] rgDispId)
        {
            // Console.WriteLine($"Name: {names[0]}");
            if (_methods_dict.TryGetValue(names[0], out var value))
            {
                rgDispId[0] = value.idx;
                return 0;
            }
            else
            {
                rgDispId[0] = DISPID_UNKNOWN;
                return DISP_E_UNKNOWNNAME;
            }
        }
        public static readonly int sizeof_VARIANT = Marshal.SizeOf<nint>() * 2 + 0x8;

        public static object? VariantToObject(nint variant, Type type)
        {
            object? obj = Marshal.GetObjectForNativeVariant(variant);
            switch (obj)
            {
                case Wrapper wrapper:
                    return wrapper._obj;
                case Int32 maybe_bool:
                    if (type == typeof(bool) && (maybe_bool == 0 || maybe_bool == 1))
                    {
                        return Convert.ToBoolean(maybe_bool);
                    }
                    return maybe_bool;
                default:
                    return obj;
            }
        }

        public static object?[] DISPPARAMS_to_objectArray(DISPPARAMS pDispParams, ParameterInfo[] parameters)
        {
            bool isVariadic = parameters.Length > 0 && parameters[parameters.Length - 1].IsDefined(typeof(ParamArrayAttribute));
            object?[] arr = new object[parameters.Length];
            int length = parameters.Length;
            if (isVariadic)
            {
                --length;
            }
            nint j = pDispParams.rgvarg + (sizeof_VARIANT * (pDispParams.cArgs - 1));
            for (int i = 0; i < length; ++i, j -= sizeof_VARIANT)
            {
                arr[i] = VariantToObject(j, parameters[i].ParameterType);
            }
            if (isVariadic)
            {
                object?[] variadicParams = new object?[pDispParams.cArgs - length];
                arr[arr.Length - 1] = variadicParams;
                Type elementType = parameters[parameters.Length - 1].ParameterType.GetElementType()!;
                for (int i = 0; i < variadicParams.Length; ++i, j -= sizeof_VARIANT)
                {
                    object? a = VariantToObject(j, elementType);
                    variadicParams[i] = a;
                }
            }
            return arr;
        }

        public static object? ConditionallyWrap(object? obj)
        {
            if (NeedsWrapping(obj))
            {
                return new Wrapper(obj);
            }
            else
            {
                return obj;
            }
        }

        public static bool NeedsWrapping([NotNullWhen(true)] object? obj)
        {
            if (obj == null)
                return false; // or true if you want to wrap null

            Type t = obj.GetType();

            // Primitive types + string + decimal + enum don't need wrapping
            if (t.IsPrimitive || t == typeof(string) || t == typeof(decimal) || t.IsEnum)
                return false;

            // Everything else should be wrapped
            return true;
        }

        public void WriteToVARIANT(object? obj, nint VarResult)
        {
            if (NeedsWrapping(obj))
            {
                Marshal.WriteInt16(VarResult, (short)VarEnum.VT_DISPATCH);
                Marshal.WriteIntPtr(VarResult, 8, Marshal.GetComInterfaceForObject(new Wrapper(obj), typeof(IDispatch)));
            }
            else
            {
                Marshal.GetNativeVariantForObject(obj, VarResult);
            }
        }

        public static bool VariantCanConvertToParam(nint variant, ParameterInfo parameterInfo)
        {
            Type type = parameterInfo.ParameterType;
            if (parameterInfo.IsDefined(typeof(ParamArrayAttribute)))
            {
                type = type.GetElementType()!;
            }
            return VariantCanConvertToType(variant, type);
        }

        public static bool VariantCanConvertToType(nint variant, Type type)
        {
            if (type == typeof(object))
            {
                return true;
            }
            // Console.WriteLine((VarEnum)Marshal.ReadInt16(variant));
            switch ((VarEnum)Marshal.ReadInt16(variant))
            {
                case VarEnum.VT_I4:
                    if (type == typeof(bool))
                    {
                        Int32 value = Marshal.ReadInt32(variant + 8);
                        if (value == 0 || value == 1)
                        {
                            return true;
                        }
                    }
                    if (type.IsEnum && type.GetEnumUnderlyingType() == typeof(Int32))
                    {
                        return true;
                    }
                    return type == typeof(Int32);
                case VarEnum.VT_I8: return type == typeof(Int64);
                case VarEnum.VT_R8: return type == typeof(double);
                case VarEnum.VT_BSTR: return type == typeof(string);
                case VarEnum.VT_DISPATCH:
                    nint pdispVal = Marshal.ReadIntPtr(variant + 8);
                    Type otherType = ((Wrapper)Marshal.GetObjectForIUnknown(pdispVal))._obj.GetType();
                    if (type == otherType) // handles value types
                    {
                        return true;
                    }
                    if (type.IsAssignableFrom(otherType))
                    {
                        return true;
                    }
                    return false;
            }
            return false;
        }

        public uint Invoke(int dispIdMember, ref Guid riid, int lcid, InvokeFlags wFlags, DISPPARAMS pDispParams, nint VarResult, nint pExcepInfo, nint puArgErr)
        {
            try
            {
                // Console.WriteLine($"dispId: {dispIdMember}");
                IEnumerable<object>? methods_or_field;
                if (dispIdMember < 0)
                {
                    if (dispIdMember == DISPID_NEWENUM
                        // && pDispParams.cArgs == 1
                        // && Marshal.ReadInt16(pDispParams.rgvarg) == (short)(VarEnum.VT_BYREF | VarEnum.VT_VARIANT) // AHK specific
                        && _obj is IEnumerable enumerable)
                    {
                        // var ptr_dest = Marshal.ReadIntPtr(pDispParams.rgvarg + 8);
                        var ptr = Marshal.GetComInterfaceForObject(new Enumerator(enumerable.GetEnumerator()), typeof(IEnumVARIANT));
                        // Marshal.WriteInt16(ptr_dest, (short)VarEnum.VT_UNKNOWN);
                        // Marshal.WriteIntPtr(ptr_dest + 8, ptr);
                        Marshal.WriteInt16(VarResult, (short)VarEnum.VT_UNKNOWN);
                        Marshal.WriteIntPtr(VarResult + 8, ptr);
                        return 0;
                    }
                    return DISP_E_MEMBERNOTFOUND;
                }
                if (dispIdMember >= _methods_list.Count || ((methods_or_field = _methods_list[dispIdMember]) == null)) // get_Item/set_Item(always dispId=0) could be null
                {
                    return DISP_E_MEMBERNOTFOUND;
                }

                switch (methods_or_field.First())
                {
                    case MethodBase method:
                        MethodBase? found_method = null;
                        ParameterInfo[]? found_parameters = null;
                        found_method = (MethodBase?)methods_or_field.FirstOrDefault(v =>
                        {
                            MethodBase method = (MethodBase)v;
                            ParameterInfo[] parameters = method.GetParameters();
                            bool isVariadic = parameters.Length > 0 && parameters[parameters.Length - 1].IsDefined(typeof(ParamArrayAttribute));
                            if (parameters.Length != pDispParams.cArgs)
                            {
                                if (!isVariadic)
                                {
                                    return false;
                                }
                                if (pDispParams.cArgs < parameters.Length - 1)
                                {
                                    return false;
                                }
                            }
                            nint j = pDispParams.rgvarg + (sizeof_VARIANT * (pDispParams.cArgs - 1));
                            for (int i = 0; i < parameters.Length; ++i, j -= sizeof_VARIANT)
                            {
                                if (!VariantCanConvertToParam(j, parameters[i]))
                                {
                                    return false;
                                }
                            }
                            if (isVariadic)
                            {
                                ParameterInfo lastParam = parameters[parameters.Length - 1];
                                for (; j >= pDispParams.rgvarg; j -= sizeof_VARIANT)
                                {
                                    if (!VariantCanConvertToParam(j, lastParam))
                                    {
                                        return false;
                                    }
                                }
                            }
                            found_parameters = parameters;
                            return true;
                        });
                        if (found_method == null)
                        {
                            return DISP_E_MEMBERNOTFOUND;
                        }
                        var args = DISPPARAMS_to_objectArray(pDispParams, found_parameters!);
                        bool isConstructor = dispIdMember == 0 && _obj is Type;
                        object? res = isConstructor ? ((ConstructorInfo)found_method).Invoke(args) : ((MethodInfo)found_method).Invoke(_obj, args);
                        if (VarResult != 0)
                        {
                            WriteToVARIANT(res, VarResult);
                        }
                        break;
                    case FieldInfo field:
                        switch (wFlags)
                        {
                            case InvokeFlags.DISPATCH_PROPERTYGET:
                                WriteToVARIANT(field.GetValue(_obj), VarResult);
                                break;
                            case InvokeFlags.DISPATCH_PROPERTYPUT:
                                if (pDispParams.cArgs != 1)
                                {
                                    return DISP_E_MEMBERNOTFOUND;
                                }
                                if (!VariantCanConvertToType(pDispParams.rgvarg, field.FieldType))
                                {
                                    return DISP_E_MEMBERNOTFOUND;
                                }
                                field.SetValue(_obj, VariantToObject(pDispParams.rgvarg, field.FieldType));
                                break;
                            default:
                                return DISP_E_MEMBERNOTFOUND;
                        }
                        break;
                }
                return 0;
            }
            catch (Exception e)
            {
                Lib.DebugWriteLine(e.ToString());
                throw;
            }
        }

        public uint GetClassInfo(out nint info)
        {
            info = Marshal.GetComInterfaceForObject(this, typeof(ITypeInfo));
            return 0;
        }

        public void GetTypeAttr(out nint ppTypeAttr)
        {
            throw new NotImplementedException();
        }

        public void GetTypeComp(out ITypeComp ppTComp)
        {
            throw new NotImplementedException();
        }

        public void GetFuncDesc(int index, out nint ppFuncDesc)
        {
            throw new NotImplementedException();
        }

        public void GetVarDesc(int index, out nint ppVarDesc)
        {
            throw new NotImplementedException();
        }

        public void GetNames(int memid, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2), Out] string[] rgBstrNames, int cMaxNames, out int pcNames)
        {
            throw new NotImplementedException();
        }

        public void GetRefTypeOfImplType(int index, out int href)
        {
            throw new NotImplementedException();
        }

        public void GetImplTypeFlags(int index, out IMPLTYPEFLAGS pImplTypeFlags)
        {
            throw new NotImplementedException();
        }

        public void GetIDsOfNames([In, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 1)] string[] rgszNames, int cNames, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1), Out] int[] pMemId)
        {
            throw new NotImplementedException();
        }

        public void Invoke([MarshalAs(UnmanagedType.IUnknown)] object pvInstance, int memid, short wFlags, ref DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, out int puArgErr)
        {
            throw new NotImplementedException();
        }

        public void GetDocumentation(int index, out string strName, nint strDocString, nint dwHelpContext, nint strHelpFile)
        {
            if (index == DISPID_UNKNOWN)
            {
                strName = _obj.GetType().Name;
                return;
            }
            throw new NotImplementedException();
        }

        public void GetDllEntry(int memid, INVOKEKIND invKind, nint pBstrDllName, nint pBstrName, nint pwOrdinal)
        {
            throw new NotImplementedException();
        }

        public void GetRefTypeInfo(int hRef, out ITypeInfo ppTI)
        {
            throw new NotImplementedException();
        }

        public void AddressOfMember(int memid, INVOKEKIND invKind, out nint ppv)
        {
            throw new NotImplementedException();
        }

        public void CreateInstance([MarshalAs(UnmanagedType.IUnknown)] object? pUnkOuter, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown), Out] out object ppvObj)
        {
            throw new NotImplementedException();
        }

        public void GetMops(int memid, out string? pBstrMops)
        {
            throw new NotImplementedException();
        }

        public void GetContainingTypeLib(out ITypeLib ppTLB, out int pIndex)
        {
            throw new NotImplementedException();
        }

        public void ReleaseTypeAttr(nint pTypeAttr)
        {
            throw new NotImplementedException();
        }

        public void ReleaseFuncDesc(nint pFuncDesc)
        {
            throw new NotImplementedException();
        }

        public void ReleaseVarDesc(nint pVarDesc)
        {
            throw new NotImplementedException();
        }
    }

    [Guid("00020401-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public interface ITypeInfo
    {
        void GetTypeAttr(out IntPtr ppTypeAttr);
        void GetTypeComp(out ITypeComp ppTComp);
        void GetFuncDesc(int index, out IntPtr ppFuncDesc);
        void GetVarDesc(int index, out IntPtr ppVarDesc);
        void GetNames(int memid, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2), Out] string[] rgBstrNames, int cMaxNames, out int pcNames);
        void GetRefTypeOfImplType(int index, out int href);
        void GetImplTypeFlags(int index, out IMPLTYPEFLAGS pImplTypeFlags);
        void GetIDsOfNames([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 1), In] string[] rgszNames, int cNames, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1), Out] int[] pMemId);
        void Invoke([MarshalAs(UnmanagedType.IUnknown)] object pvInstance, int memid, short wFlags, ref DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr);
        void GetDocumentation(int index, out string strName, nint strDocString, nint dwHelpContext, nint strHelpFile);
        void GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal);
        void GetRefTypeInfo(int hRef, out ITypeInfo ppTI);
        void AddressOfMember(int memid, INVOKEKIND invKind, out IntPtr ppv);
        void CreateInstance([MarshalAs(UnmanagedType.IUnknown)] object? pUnkOuter, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown), Out] out object ppvObj);
        void GetMops(int memid, out string? pBstrMops);
        void GetContainingTypeLib(out ITypeLib ppTLB, out int pIndex);
        [PreserveSig]
        void ReleaseTypeAttr(IntPtr pTypeAttr);
        [PreserveSig]
        void ReleaseFuncDesc(IntPtr pFuncDesc);
        [PreserveSig]
        void ReleaseVarDesc(IntPtr pVarDesc);
    }

    [Guid("00020404-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public interface IEnumVARIANT
    {
        [PreserveSig]
        int Next(int celt, nint rgVar, nint pceltFetched);

        [PreserveSig]
        int Skip(int celt);

        [PreserveSig]
        int Reset();

        IEnumVARIANT Clone();
    }

    public class Enumerator : IEnumVARIANT
    {
        public Enumerator(IEnumerator enumerator)
        {
            _enumerator = enumerator;
            _idx = 0;
            _saved = [];
        }

        private IEnumerator _enumerator;
        private int _idx;
        private List<object?> _saved;

        public Enumerator(Enumerator that)
        {
            _enumerator = that._enumerator;
            _idx = that._idx;
            _saved = that._saved;
        }

        public int Next(int celt, nint rgVar, nint pceltFetched)
        {
            int i = 0;
            for (int j = 0; i < celt; ++i, j += Wrapper.sizeof_VARIANT)
            {
                if (_idx >= _saved.Count)
                {
                    if (_enumerator.MoveNext())
                    {
                        _saved.Add(Wrapper.ConditionallyWrap(_enumerator.Current));
                    }
                    else
                    {
                        break;
                    }
                }
                object? res = _saved[_idx];
                nint VarResult = rgVar + j;
                if (res is Wrapper wrapper)
                {
                    Marshal.WriteInt16(VarResult, (short)VarEnum.VT_DISPATCH);
                    Marshal.WriteIntPtr(VarResult, 8, Marshal.GetComInterfaceForObject(wrapper, typeof(IDispatch)));
                }
                else
                {
                    Marshal.GetNativeVariantForObject(res, VarResult);
                }
                ++_idx;
            }
            if (pceltFetched != 0)
            {
                Marshal.WriteInt32(pceltFetched, i);
            }
            if (i == celt)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        public int Skip(int celt)
        {
            // implementation taken from Next(), but trimmed
            int i = 0;
            for (; i < celt; ++i)
            {
                if (_idx >= _saved.Count)
                {
                    if (_enumerator.MoveNext())
                    {
                        _saved.Add(Wrapper.ConditionallyWrap(_enumerator.Current));
                    }
                    else
                    {
                        break;
                    }
                }
                ++_idx;
            }
            if (i == celt)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        public int Reset()
        {
            _idx = 0;
            return 0;
        }

        public IEnumVARIANT Clone()
        {
            return new Enumerator(this);
        }
    }

}
