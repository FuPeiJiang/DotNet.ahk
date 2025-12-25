#Requires AutoHotkey v2.0

DotNet_LoadLibrary(AssemblyName) {
    return DotNet.LoadAssembly(AssemblyName)
}

; class DotNet_Change_Path {
;     static _ := DotNet_LoadLibrary.AHK_DotNet_Interop_Path := "C:\custom\AHK_Dotnet_Interop.dll"
; }

class DotNet {

    static __New() {
        toSort:=[]
        Loop Files "C:\Program Files\dotnet\shared\Microsoft.NETCore.App\*", "D" {
            ; semver: major.minor.patch-prerelease
            RegExMatch(A_LoopFileName, "^(\d+)\.(\d+)\.(\d+)(?:-(.+))?$", &OutputVar)
            if (!OutputVar) {
                continue
            }
            toSort.Push(OutputVar)
        }
        if (!toSort.Length) {
            throw Error("No .NET version found in C:\Program Files\dotnet\shared\Microsoft.NETCore.App")
        }
        InsertionSort(toSort, semver_cmp)
        latest_version := toSort[toSort.Length]
        NETCore_path := "C:\Program Files\dotnet\shared\Microsoft.NETCore.App\" latest_version.0
        coreclr_fullpath := NETCore_path "\coreclr.dll"
        coreclr_hModule := DllCall("LoadLibraryW", "WStr", coreclr_fullpath, "Ptr")
        SplitPath A_AhkPath,, &exeDir
        TRUSTED_PLATFORM_ASSEMBLIES := ""
        Loop Files NETCore_path "\*.dll", "F" {
            if (A_Index > 1) {
                TRUSTED_PLATFORM_ASSEMBLIES .= ";"
            }
            TRUSTED_PLATFORM_ASSEMBLIES .= A_LoopFileFullPath
        }
        propertyKeys := Buffer(A_PtrSize)
        propertyValues := Buffer(A_PtrSize)
        TRUSTED_PLATFORM_ASSEMBLIES_UTF8 := UTF8("TRUSTED_PLATFORM_ASSEMBLIES")
        TRUSTED_PLATFORM_ASSEMBLIES_LIST := UTF8(TRUSTED_PLATFORM_ASSEMBLIES)
        NumPut("Ptr", TRUSTED_PLATFORM_ASSEMBLIES_UTF8.Ptr, propertyKeys)
        NumPut("Ptr", TRUSTED_PLATFORM_ASSEMBLIES_LIST.Ptr, propertyValues)
        ; https://github.com/dotnet/runtime/blob/main/src/coreclr/hosts/inc/coreclrhost.h
        DllCall("coreclr\coreclr_initialize", "Ptr", UTF8(exeDir), "AStr", "AutoHotkeyHost", "Int", 1, "Ptr", propertyKeys, "Ptr", propertyValues, "Ptr*", &hostHandle:=0, "Uint*", &domainId:=0)
        DllCall("coreclr\coreclr_create_delegate", "Ptr", hostHandle, "Uint", domainId, "AStr", "System.Private.CoreLib", "AStr", "Internal.Runtime.InteropServices.ComponentActivator", "AStr", "LoadAssemblyBytes", "Ptr*", &load_assembly_bytes:=0)
        DllCall("coreclr\coreclr_create_delegate", "Ptr", hostHandle, "Uint", domainId, "AStr", "System.Private.CoreLib", "AStr", "Internal.Runtime.InteropServices.ComponentActivator", "AStr", "GetFunctionPointer", "Ptr*", &get_function_pointer:=0)
        ; https://github.com/dotnet/runtime/blob/main/src/libraries/System.Private.CoreLib/src/Internal/Runtime/InteropServices/ComponentActivator.cs
        ; https://github.com/dotnet/runtime/blob/main/src/native/corehost/coreclr_delegates.h
        load_assembly(NETCore_path "\System.Runtime.dll")
        load_assembly(NETCore_path "\System.Runtime.InteropServices.dll")
        load_assembly(NETCore_path "\System.Console.dll")
        load_assembly(NETCore_path "\System.Threading.dll")
        load_assembly(NETCore_path "\System.Text.Encoding.Extensions.dll")
        load_assembly(NETCore_path "\System.Linq.dll")
        load_assembly(NETCore_path "\System.Collections.dll")
        load_assembly(DotNet_LoadLibrary.HasOwnProp("AHK_DotNet_Interop_Path") ? DotNet_LoadLibrary.AHK_DotNet_Interop_Path : A_LineFile "\..\AHK_DotNet_Interop.dll")
        load_assembly(path, symbol := false) {
            ; OutputDebug "loading assembly: " path "`n"
            try {
                if (symbol) {
                    assembly := FileRead(path, "RAW")
                    symbol := FileRead(path, "RAW")
                    hr := DllCall(load_assembly_bytes, "Ptr", assembly, "Ptr", assembly.Size, "Ptr", symbol, "Ptr", symbol.Size, "Ptr", 0, "Ptr", 0)
                } else {
                    assembly := FileRead(path, "RAW")
                    hr := DllCall(load_assembly_bytes, "Ptr", assembly, "Ptr", assembly.Size, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0)
                }
            } catch OsError as e {
                if (e.What == "FileRead") {
                    MsgBox "Failed to load assembly:`n`n" path "`n`n" e.Message
                } else {
                    throw e
                }
            }
            ; OutputDebug "hr: " hr "`n"
        }

        DllCall("coreclr\coreclr_create_delegate", "Ptr", hostHandle, "Uint", domainId, "AStr", "AHK_DotNet_Interop", "AStr", "AHK_DotNet_Interop.Lib", "AStr", "GetClass", "Ptr*", &GetClass_delegate:=0)
        DotNet.GetClass_delegate := GetClass_delegate
        DllCall("coreclr\coreclr_create_delegate", "Ptr", hostHandle, "Uint", domainId, "AStr", "AHK_DotNet_Interop", "AStr", "AHK_DotNet_Interop.Lib", "AStr", "LoadAssembly", "Ptr*", &LoadAssembly_delegate:=0)
        DotNet.LoadAssembly_delegate := LoadAssembly_delegate

        UTF8(str) {
            ; StrPut: In 2-parameter mode, this function returns the required buffer size in bytes,
            ; including space for the null-terminator.
            buf := Buffer(StrPut(str, "UTF-8"))
            StrPut(str, buf, "UTF-8")
            return buf
        }

        semver_cmp(a, b) {
            major := a.1 - b.1
            if (major) {
                return major
            }
            minor := a.2 - b.2
            if (minor) {
                return minor
            }
            patch := a.3 - b.3
            if (patch) {
                return patch
            }
            ; no pre-release is higher
            if (!a.4) {
                return 1
            }
            if (!b.4) {
                return -1
            }
            ; string compare (bad)
            return a > b ? 1 : -1 ; cannot be equal, so else is a < b
        }

        InsertionSort(A, cmp) {
            ; https://en.wikipedia.org/wiki/Insertion_sort#Algorithm
            ; i ← 1
            ; while i < length(A)
            ;     x ← A[i]
            ;     j ← i
            ;     while j > 0 and A[j-1] > x
            ;         A[j] ← A[j-1]
            ;         j ← j - 1
            ;     end while
            ;     A[j] ← x
            ;     i ← i + 1
            ; end while
            i := 2
            while (i <= A.Length) {
                x := A[i]
                j := i
                while (j > 1 && cmp(A[j-1], x) > 0) {
                    A[j] := A[j - 1]
                    j := j - 1
                }
                A[j] := x
                i := i + 1
            }
            return A
        }
    }

    static LoadAssembly(path) {
        DllCall(DotNet.LoadAssembly_delegate, "WStr", path, "Ptr*", IDisPatch:=ComValue(9, 0))
        return IDisPatch
    }

    static using(FullName) {
        DllCall(DotNet.GetClass_delegate, "AStr", FullName, "Ptr*", IDisPatch:=ComValue(9, 0))
        return IDisPatch
    }
}

; Console := DotNet.using("System.Console")
; Console.WriteLine("Hello from C#")

; File_ := DotNet.using("System.IO.File")
; Console.WriteLine(File_.ReadAllText(A_LineFile))
