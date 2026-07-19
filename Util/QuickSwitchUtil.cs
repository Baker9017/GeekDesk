using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;

namespace GeekDesk.Util
{
    /// <summary>
    /// Listary 风格 "Quick Switch (文件夹跟随)" 核心实现。
    ///
    /// 职责:
    ///   1. 通过 SetWinEventHook 监听 #32770 对话框的创建/显示/销毁
    ///   2. 通过 Shell COM (IShellWindows) + UI Automation 读取最近资源管理器窗口的当前路径
    ///   3. 把路径注入到对话框地址栏 Edit 控件, 模拟回车触发导航
    ///   4. 对每个被跟踪的对话框, 由调用方轮询持续同步
    /// </summary>
    public static class QuickSwitchUtil
    {
        #region Shell COM (读取 Explorer 路径)

        // Win10/Win11 上 UI Automation 读取地址栏 Edit 经常拿不到 ValuePattern,
        // 而 Shell COM (IShellWindows -> IWebBrowserApp.LocationURL) 是稳定的备用通道.
        // 我们直接走 IDispatch 调度, 不引用 SHDocVw 程序集.

        [DllImport("ole32.dll")]
        private static extern int CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext, ref Guid riid, out IntPtr ppv);

        [DllImport("oleaut32.dll")]
        private static extern int VariantClear(IntPtr pVarg);

        // MSAA (Microsoft Active Accessibility) 备用路径.
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, out IntPtr ppvObject);

        private static readonly Guid IID_IAccessible = new Guid("618736E0-3C3D-11CF-810C-02803C1174A1");
        // IWebBrowserApp - 直接请求这个接口, 拿到后用 IDispatch 调度 LocationURL.
        private static readonly Guid IID_IWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
        // IServiceProvider - 用于在 IE/Explorer 浏览器对象上查询子接口.
        private static readonly Guid IID_IServiceProvider = new Guid("6D5140C1-7436-11CE-8034-00AA006009FA");
        // IShellBrowser - 用来枚举当前 tab.
        private static readonly Guid IID_IShellBrowser = new Guid("C08AFD90-F2A1-11D1-8455-00A0C91F3880");
        // IID_IHlinkFrame (备选) etc.
        // OBJID_WINDOW 用类里已经声明的 public const

        private static readonly Guid CLSID_ShellWindows = new Guid("9BA05972-F6A8-11CF-A442-00A0C90A8F39");
        // IShellWindows 是 dispinterface, 直接通过 IID_IShellWindows (继承 IDispatch) 创建会返回 E_NOINTERFACE.
        // 改用 IID_IDispatch 创建, 拿到的是 dispatch 指针, 用 GetIDsOfNames/Invoke 调度.
        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
        private const uint CLSCTX_LOCAL_SERVER = 0x4;
        private const uint CLSCTX_INPROC_SERVER = 0x1;
        private const uint CLSCTX_ALL = 0x17;

        [Flags]
        private enum DispatchFlags : ushort
        {
            DISPATCH_METHOD = 0x1,
            DISPATCH_PROPERTYGET = 0x2,
            DISPATCH_PROPERTYPUT = 0x4,
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPPARAMS
        {
            public IntPtr rgvarg;
            public IntPtr rgdispidNamedArgs;
            public uint cArgs;
            public uint cNamedArgs;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct EXCEPINFO
        {
            public ushort wCode;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            [MarshalAs(UnmanagedType.BStr)] public string bstrSource;
            [MarshalAs(UnmanagedType.BStr)] public string bstrDescription;
            [MarshalAs(UnmanagedType.BStr)] public string bstrHelpFile;
            public uint dwHelpContext;
            public IntPtr pfnDeferredFillIn;
            public IntPtr scode;
        }

        // 走 IDispatch vtable (GetIDsOfNames / Invoke)
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00020400-0000-0000-C000-000000000046")]
        private interface IDispatch
        {
            [PreserveSig] int GetTypeInfoCount(out uint pctinfo);
            [PreserveSig] int GetTypeInfo(uint iTInfo, int lcid, out IntPtr ppTInfo);
            [PreserveSig] int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, uint cNames, int lcid, [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig] int Invoke(int dispIdMember, ref Guid riid, int lcid, DispatchFlags wFlags, ref DISPPARAMS pDispParams, out IntPtr pVarResult, ref EXCEPINFO pExcepInfo, out uint puArgErr);
        }

        /// <summary>
        /// 通过 Shell COM (IShellWindows) 获取某个指定 Explorer 窗口 (targetHwnd!=0)
        /// 或任意可见 Explorer 窗口 (targetHwnd==0) 的当前路径。
        /// 必须在 STA 线程上调用。失败返回 null。
        /// </summary>
        public static string GetExplorerPathViaShell(IntPtr targetHwnd)
        {
            IntPtr shellWindowsPtr = IntPtr.Zero;
            try
            {
                Guid clsidShell = CLSID_ShellWindows;
                Guid iidUnknown = IID_IUnknown;
                int hr = CoCreateInstance(ref clsidShell, IntPtr.Zero, CLSCTX_ALL,
                    ref iidUnknown, out shellWindowsPtr);
                LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: CoCreateInstance hr=0x" + hr.ToString("X") + " ptr=" + shellWindowsPtr.ToString("X"));
                if (hr != 0 || shellWindowsPtr == IntPtr.Zero)
                {
                    return null;
                }

                IntPtr dispatchPtr;
                Guid iidDispatch = IID_IDispatch;
                int hrQi = Marshal.QueryInterface(shellWindowsPtr, ref iidDispatch, out dispatchPtr);
                Marshal.Release(shellWindowsPtr);
                LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: QI(IDispatch) hr=0x" + hrQi.ToString("X") + " ptr=" + dispatchPtr.ToString("X"));
                if (hrQi != 0 || dispatchPtr == IntPtr.Zero)
                {
                    return null;
                }

                var shellWindows = (IDispatch)Marshal.GetObjectForIUnknown(dispatchPtr);
                Marshal.Release(dispatchPtr);
                if (shellWindows == null) return null;

                try
                {
                    int[] dispIds = new int[1];
                    Guid g = IID_IUnknown;

                    // Count
                    int hrCountNames = shellWindows.GetIDsOfNames(ref g, new string[] { "Count" }, 1, 0, dispIds);
                    if (hrCountNames != 0)
                    {
                        LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: GetIDsOfNames(Count) hr=0x" + hrCountNames.ToString("X"));
                        return null;
                    }
                    IntPtr varResult;
                    EXCEPINFO ex = new EXCEPINFO();
                    uint argErr;
                    DISPPARAMS p = new DISPPARAMS();
                    int hrCount = shellWindows.Invoke(dispIds[0], ref g, 0, DispatchFlags.DISPATCH_PROPERTYGET,
                            ref p, out varResult, ref ex, out argErr);
                    LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: Invoke(Count) hr=0x" + hrCount.ToString("X") + " var=" + (varResult == IntPtr.Zero ? "0" : "set"));
                    if (hrCount != 0 || varResult == IntPtr.Zero)
                    {
                        LogUtil.WriteQuickSwitchLog("  excep wCode=0x" + ex.wCode.ToString("X") + " source=" + (ex.bstrSource ?? "null") + " desc=" + (ex.bstrDescription ?? "null"));
                        return null;
                    }
                    int count = Marshal.ReadInt32(varResult, 8); // variant VT_I4 @ offset 8
                    VariantClear(varResult);

                    LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: ShellWindows.Count = " + count);
                    if (count <= 0) return null;

                    if (shellWindows.GetIDsOfNames(ref g, new string[] { "Item" }, 1, 0, dispIds) != 0)
                        return null;

                    for (int i = 0; i < count; i++)
                    {
                        IntPtr itemResult = IntPtr.Zero;
                        IntPtr vIdx = IntPtr.Zero;
                        try
                        {
                            vIdx = Marshal.AllocHGlobal(16);
                            Marshal.WriteInt16(vIdx, 0, 3);   // VT_I4
                            Marshal.WriteInt16(vIdx, 2, 0);
                            Marshal.WriteInt32(vIdx, 8, i);
                            p = new DISPPARAMS { rgvarg = vIdx, cArgs = 1 };
                            if (shellWindows.Invoke(dispIds[0], ref g, 0, DispatchFlags.DISPATCH_METHOD,
                                    ref p, out itemResult, ref ex, out argErr) != 0 || itemResult == IntPtr.Zero)
                                continue;

                            var item = (IDispatch)Marshal.GetObjectForIUnknown(itemResult);
                            if (item == null) { Marshal.Release(itemResult); continue; }

                            try
                            {
                                // HWND
                                int[] subIds = new int[1];
                                if (item.GetIDsOfNames(ref g, new string[] { "HWND" }, 1, 0, subIds) != 0) continue;
                                IntPtr hwVar;
                                DISPPARAMS pp = new DISPPARAMS();
                                if (item.Invoke(subIds[0], ref g, 0, DispatchFlags.DISPATCH_PROPERTYGET,
                                        ref pp, out hwVar, ref ex, out argErr) != 0 || hwVar == IntPtr.Zero)
                                    continue;
                                IntPtr shellHwnd = new IntPtr(Marshal.ReadInt32(hwVar, 8));
                                VariantClear(hwVar);

                                if (targetHwnd != IntPtr.Zero)
                                {
                                    if (shellHwnd != targetHwnd) continue;
                                }
                                else
                                {
                                    if (shellHwnd == IntPtr.Zero || !IsWindow(shellHwnd) || !IsWindowVisible(shellHwnd))
                                        continue;
                                }

                                // LocationURL
                                if (item.GetIDsOfNames(ref g, new string[] { "LocationURL" }, 1, 0, subIds) != 0) continue;
                                IntPtr urlVar;
                                if (item.Invoke(subIds[0], ref g, 0, DispatchFlags.DISPATCH_PROPERTYGET,
                                        ref pp, out urlVar, ref ex, out argErr) != 0 || urlVar == IntPtr.Zero)
                                    continue;
                                string url = (string)Marshal.GetObjectForNativeVariant(urlVar);
                                VariantClear(urlVar);

                                if (string.IsNullOrEmpty(url)) continue;
                                string path = UrlToPath(url);
                                if (!string.IsNullOrEmpty(path) && IsValidPath(path))
                                {
                                    return path;
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(item);
                            }
                        }
                        catch (Exception ex2)
                        {
                            LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell item EX: " + ex2.Message);
                        }
                        finally
                        {
                            if (vIdx != IntPtr.Zero) Marshal.FreeHGlobal(vIdx);
                            if (itemResult != IntPtr.Zero)
                            {
                                try { Marshal.Release(itemResult); } catch { }
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(shellWindows);
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell EX: " + ex.Message);
            }
            return null;
        }

        #endregion

        #region MSAA / WM_GETOBJECT 读取 Explorer 地址栏

        // WM_GETOBJECT + LresultFromObject 是 OleAcc 提供的一种获取 accessibility 对象的机制,
        // 可以在 AccessibleObjectFromWindow 失败时作为备选.
        private const int WM_GETOBJECT = 0x003D;
        private const int WM_COPYDATA = 0x004A;

        // IDispatch 的 DISPID for common members:
        //   DISPID_NEWENUM = -4 (newenum)
        //   DISPID_VALUE   = 0  (default property)
        //   DISPID_COUNT   = 1  (count)
        private const int DISPID_NEWENUM = -4;
        private const int DISPID_VALUE = 0;

        /// <summary>
        /// 方式1: AccessibleObjectFromWindow → IWebBrowserApp → LocationURL.
        /// 方式2: WM_GETOBJECT → IDispatch → 遍历子窗口找 LocationURL.
        /// 必须 STA 线程。
        /// </summary>
        public static string GetExplorerPathViaAccessible(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd)) return null;

            uint pid = 0;
            GetWindowThreadProcessId(hwnd, out pid);
            string cls = GetClassNameString(hwnd);

            // 方式1: AccessibleObjectFromWindow (和之前一样, 保留以防某些系统可用)
            IntPtr pDisp = IntPtr.Zero;
            try
            {
                Guid iidWb = IID_IWebBrowserApp;
                int hr = AccessibleObjectFromWindow(hwnd, OBJID_WINDOW, ref iidWb, out pDisp);
                if (hr == 0 && pDisp != IntPtr.Zero)
                {
                    string path = TryReadLocationUrl(pDisp);
                    Marshal.Release(pDisp);
                    pDisp = IntPtr.Zero;
                    if (!string.IsNullOrEmpty(path)) return path;
                }
            }
            catch { }
            finally { if (pDisp != IntPtr.Zero) { Marshal.Release(pDisp); pDisp = IntPtr.Zero; } }

            // 方式2: WM_GETOBJECT → 直接取窗口的 IDispatch (LresultFromObject).
            //    发 WM_GETOBJECT 到目标窗口, OLEACC 返回一个 LRESULT, 用 ObjectFromLresult 解出 IDispatch.
            pDisp = TryGetDispatchViaWmGetObject(hwnd);
            if (pDisp != IntPtr.Zero)
            {
                try
                {
                    string path = TryReadLocationUrl(pDisp);
                    if (!string.IsNullOrEmpty(path)) return path;

                    // 如果当前 IDispatch 本身不是浏览器对象, 遍历 Item() / Children.
                    path = TryEnumDispatchChildren(pDisp, depth: 0);
                    if (!string.IsNullOrEmpty(path)) return path;
                }
                finally { Marshal.Release(pDisp); pDisp = IntPtr.Zero; }
            }

            // 方式3: UI Automation (最底层保障).
            try
            {
                var ae = AutomationElement.FromHandle(hwnd);
                if (ae != null)
                {
                    string path = TryReadPathFromUIAutomation(ae);
                    if (!string.IsNullOrEmpty(path)) return path;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 通过 SendMessage(WM_GETOBJECT, 0, lParam) + ObjectFromLresult 获取窗口的 IDispatch.
        /// 这是 AccessibleObjectFromWindow 的底层实现, 在某些情况下更可靠.
        /// </summary>
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(
            IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam,
            uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

        private const uint SMTO_ABORTIFHUNG = 0x0002;

        [DllImport("oleacc.dll")]
        private static extern int ObjectFromLresult(IntPtr lResult, ref Guid riid, IntPtr wParam, out IntPtr ppvObject);

        private static IntPtr TryGetDispatchViaWmGetObject(IntPtr hwnd)
        {
            IntPtr lpdwResult;
            IntPtr lResult = SendMessageTimeout(hwnd, WM_GETOBJECT, IntPtr.Zero, (IntPtr)OBJID_WINDOW,
                SMTO_ABORTIFHUNG, 500, out lpdwResult);
            if (lResult == IntPtr.Zero || lResult == new IntPtr(-1)) return IntPtr.Zero;

            Guid iid = IID_IDispatch;
            IntPtr pDisp;
            int hr = ObjectFromLresult(lResult, ref iid, IntPtr.Zero, out pDisp);
            if (hr == 0 && pDisp != IntPtr.Zero) return pDisp;

            // 备选: 请求 IAccessible
            iid = IID_IAccessible;
            hr = ObjectFromLresult(lResult, ref iid, IntPtr.Zero, out pDisp);
            if (hr == 0 && pDisp != IntPtr.Zero)
            {
                // 把 IAccessible 转成 IDispatch
                IntPtr pUnk = IntPtr.Zero;
                try
                {
                    Guid iidUnk = IID_IUnknown;
                    hr = Marshal.QueryInterface(pDisp, ref iidUnk, out pUnk);
                    if (hr == 0 && pUnk != IntPtr.Zero)
                    {
                        var obj = Marshal.GetObjectForIUnknown(pUnk);
                        if (obj != null)
                        {
                            pUnk = IntPtr.Zero; // obj 现在持有引用
                            Marshal.ReleaseComObject(obj);
                            return pDisp; // 返回 IAccessible 指针, TryReadLocationUrl 会处理
                        }
                    }
                }
                finally
                {
                    if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
                }
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// 遍历 IDispatch 集合的子成员 (Item() / _NewEnum), 找 LocationURL.
        /// </summary>
        private static string TryEnumDispatchChildren(IntPtr pDisp, int depth)
        {
            if (pDisp == IntPtr.Zero || depth > 4) return null;
            try
            {
                var obj = Marshal.GetObjectForIUnknown(pDisp);
                if (obj == null) return null;
                try
                {
                    Type t = obj.GetType();

                    // 先试试 LocationURL
                    try
                    {
                        object val = t.InvokeMember("LocationURL", System.Reflection.BindingFlags.GetProperty, null, obj, null);
                        if (val is string s && !string.IsNullOrEmpty(s))
                        {
                            string path = UrlToPath(s);
                            if (!string.IsNullOrEmpty(path)) return path;
                        }
                    }
                    catch { }

                    // 遍历 Item(index) 或 _NewEnum
                    try
                    {
                        // 先尝试 _NewEnum (DISPID = -4)
                        object newEnum = t.InvokeMember("_NewEnum",
                            System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.GetField,
                            null, obj, null);
                        if (newEnum != null)
                        {
                            // IEnumVARIANT 遍历
                            string path = TryEnumViaIEnumVariant(newEnum, depth + 1);
                            if (!string.IsNullOrEmpty(path)) return path;
                        }
                    }
                    catch { }

                    // 尝试 Item(index) for index = 0..15
                    for (int i = 0; i < 16; i++)
                    {
                        try
                        {
                            object item = t.InvokeMember("Item",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, obj, new object[] { i });
                            if (item != null)
                            {
                                IntPtr pItem = Marshal.GetIUnknownForObject(item);
                                if (pItem != IntPtr.Zero)
                                {
                                    try
                                    {
                                        string path = TryReadLocationUrl(pItem);
                                        if (!string.IsNullOrEmpty(path)) return path;
                                    }
                                    finally { Marshal.Release(pItem); }
                                }
                                Marshal.ReleaseComObject(item);
                            }
                        }
                        catch { }
                    }
                }
                finally { Marshal.ReleaseComObject(obj); }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 对 IEnumVARIANT 调用 Next() 遍历子元素, 找 LocationURL.
        /// </summary>
        private static string TryEnumViaIEnumVariant(object enumObj, int depth)
        {
            if (enumObj == null || depth > 3) return null;
            try
            {
                Type t = enumObj.GetType();
                // IEnumVARIANT::Next(1, out var, out fetched)
                object[] args = new object[3];
                args[0] = 1; // celt
                int fetched = 0;

                for (int i = 0; i < 20; i++)
                {
                    args[1] = null;
                    args[2] = 0;
                    try
                    {
                        t.InvokeMember("Next",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null, enumObj, args);
                        fetched = (int)args[2];
                    }
                    catch
                    {
                        // Next 可能用 DISPID 0 (Value) 作为方法名
                        try
                        {
                            object r = t.InvokeMember("Next",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, enumObj,
                                new object[] { 1 });
                            if (r != null && r.GetType().IsArray)
                            {
                                var arr = (Array)r;
                                if (arr.Length >= 1) args[1] = arr.GetValue(0);
                                fetched = arr.Length;
                            }
                        }
                        catch { break; }
                    }

                    if (fetched <= 0 || args[1] == null) break;

                    // 获取 IDispatch / IWebBrowserApp 指针
                    IntPtr pItem = IntPtr.Zero;
                    try
                    {
                        if (args[1] is DispatchWrapper dw)
                        {
                            pItem = Marshal.GetIUnknownForObject(dw.WrappedObject);
                        }
                        else
                        {
                            pItem = Marshal.GetIUnknownForObject(args[1]);
                        }

                        if (pItem != IntPtr.Zero)
                        {
                            string path = TryReadLocationUrl(pItem);
                            if (!string.IsNullOrEmpty(path)) return path;
                        }
                    }
                    catch { }
                    finally { if (pItem != IntPtr.Zero) Marshal.Release(pItem); }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// UI Automation 最底层保障: 找 Edit/ComboBox 控件的 ValuePattern.
        /// </summary>
        private static string TryReadPathFromUIAutomation(AutomationElement ae)
        {
            try
            {
                var edits = ae.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                foreach (AutomationElement edit in edits)
                {
                    try
                    {
                        if (edit.Current.Name != null && IsValidPath(edit.Current.Name))
                            return edit.Current.Name;
                        var vp = edit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        if (vp != null && !string.IsNullOrEmpty(vp.Current.Value) && IsValidPath(vp.Current.Value))
                            return vp.Current.Value;
                    }
                    catch { }
                }

                var combos = ae.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox));
                foreach (AutomationElement combo in combos)
                {
                    try
                    {
                        var vp = combo.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        if (vp != null && !string.IsNullOrEmpty(vp.Current.Value) && IsValidPath(vp.Current.Value))
                            return vp.Current.Value;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 给定一个 IDispatch 指针, 通过 GetIDsOfNames("LocationURL") / Invoke 拿到 URL 字符串。
        /// </summary>
        private static string TryReadLocationUrl(IntPtr pDisp)
        {
            if (pDisp == IntPtr.Zero) return null;
            try
            {
                // Marshal.GetObjectForIUnknown 拿到 RCW, 然后用反射 InvokeMember "LocationURL".
                var obj = Marshal.GetObjectForIUnknown(pDisp);
                if (obj == null) return null;
                try
                {
                    object val = obj.GetType().InvokeMember("LocationURL",
                        System.Reflection.BindingFlags.GetProperty,
                        null, obj, null);
                    if (val is string s && !string.IsNullOrEmpty(s))
                    {
                        string path = UrlToPath(s);
                        if (!string.IsNullOrEmpty(path) && IsValidPath(path)) return path;
                    }
                    else
                    {
                        LogUtil.WriteQuickSwitchLog("TryReadLocationUrl: LocationURL val=" + (val ?? (object)"null") + " url=" + (val as string ?? ""));
                    }
                }
                catch (Exception ex2)
                {
                    LogUtil.WriteQuickSwitchLog("TryReadLocationUrl: InvokeMember LocationURL EX: " + ex2.Message);
                }
                finally
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("TryReadLocationUrl EX: " + ex.Message);
            }
            return null;
        }


        #endregion

        private static string UrlToPath(string url)
        {
            if (string.IsNullOrEmpty(url)) return null;
            try
            {
                string p;
                if (url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                {
                    p = url.Substring("file:///".Length);
                }
                else if (url.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
                {
                    p = url.Substring("file:".Length);
                }
                else
                {
                    return null;
                }

                // 去掉可能的 query / fragment
                int hash = p.IndexOf('#');
                if (hash >= 0) p = p.Substring(0, hash);
                int qm = p.IndexOf('?');
                if (qm >= 0) p = p.Substring(0, qm);

                // unescape + 转反斜杠
                p = Uri.UnescapeDataString(p);
                p = p.Replace('/', '\\');

                // 去掉 leading "\\" (UNC 不需要多一个反斜杠)
                if (p.StartsWith("\\\\")) p = p.Substring(1);

                return p;
            }
            catch
            {
                return null;
            }
        }

        #region Win32 常量

        public const uint EVENT_OBJECT_CREATE = 0x8000;
        public const uint EVENT_OBJECT_SHOW = 0x8002;
        public const uint EVENT_OBJECT_DESTROY = 0x8001;
        public const uint EVENT_OBJECT_HIDE = 0x8003;
        public const uint EVENT_OBJECT_FOCUS = 0x8005;
        public const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B;
        public const uint EVENT_SYSTEM_FOREGROUND = 0x0003;

        public const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        public const uint WINEVENT_SKIPOWNPROCESS = 0x0002;

        public const int OBJID_WINDOW = 0x00000000;

        private const int WM_SETTEXT = 0x000C;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int VK_RETURN = 0x0D;

        #endregion

        #region Win32 API

        public delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        #endregion

        #region 状态 (线程安全)

        public struct DialogEvent
        {
            public uint EventType;
            public IntPtr Hwnd;
        }

        private static int _geekDeskPid;
        private static long _lastExplorerPathTimestamp;
        private static string _lastExplorerPath;
        private static readonly object _lastExplorerLock = new object();

        /// <summary>
        /// 队列: 新出现/销毁的对话框 HWND, 由 QuickSwitchThread 的工作线程消费。
        /// </summary>
        public static readonly ConcurrentQueue<DialogEvent> DialogEvents = new ConcurrentQueue<DialogEvent>();

        /// <summary>
        /// 当前正在被持续同步的对话框列表 (HWND -> 路径缓存)。
        /// </summary>
        public static readonly ConcurrentDictionary<IntPtr, string> TrackedDialogs = new ConcurrentDictionary<IntPtr, string>();

        #endregion

        /// <summary>
        /// 初始化时记录 GeekDesk 自身 PID, 用于过滤自身对话框。
        /// </summary>
        public static void InitSelfPid()
        {
            try
            {
                using (var p = System.Diagnostics.Process.GetCurrentProcess())
                {
                    _geekDeskPid = p.Id;
                }
            }
            catch
            {
                _geekDeskPid = 0;
            }
        }

        public static int GeekDeskPid => _geekDeskPid;

        public static string GetLastExplorerPath()
        {
            lock (_lastExplorerLock)
            {
                return _lastExplorerPath;
            }
        }

        public static long GetLastExplorerTimestamp()
        {
            lock (_lastExplorerLock)
            {
                return _lastExplorerPathTimestamp;
            }
        }

        public static void SetLastExplorerPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            lock (_lastExplorerLock)
            {
                _lastExplorerPath = path;
                _lastExplorerPathTimestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            }
        }

        /// <summary>
        /// 获取当前前台 Explorer 窗口的路径, 通过 UI Automation 读取地址栏。
        /// 必须从 STA 线程调用。
        /// </summary>
        public static string GetCurrentExplorerPath()
        {
            try
            {
                IntPtr fg = GetForegroundWindow();
                if (fg == IntPtr.Zero) return null;

                string cls = GetClassNameString(fg);
                if (!IsExplorerClass(cls)) return null;

                return GetPathFromExplorerWindow(fg);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 遍历所有可见 Explorer 窗口, 取首个路径。
        /// 备用方案, 必须从 STA 线程调用。
        /// 优先走 Shell COM, 失败时 fallback 到 UI Automation 枚举.
        /// </summary>
        public static string GetAnyExplorerPath()
        {
            try
            {
                string shellPath = GetExplorerPathViaShell(IntPtr.Zero);
                if (!string.IsNullOrEmpty(shellPath) && IsValidPath(shellPath))
                {
                    return shellPath;
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("GetAnyExplorerPath Shell EX: " + ex.Message);
            }

            // MSAA fallback: 遍历所有可见 Explorer 窗口
            _latestExplorerResult = null;
            try
            {
                EnumWindows(EnumWindowsImpl, IntPtr.Zero);
            }
            catch
            {
                // ignored
            }
            return _latestExplorerResult;
        }

        // EnumWindows 回调不能用 lambda (lambda 在 native interop 上不稳定)
        private static bool EnumWindowsImpl(IntPtr hwnd, IntPtr lParam)
        {
            if (!IsWindowVisible(hwnd)) return true;
            string cls = GetClassNameString(hwnd);
            if (!IsExplorerClass(cls)) return true;

            string p = GetPathFromExplorerWindow(hwnd);
            if (!string.IsNullOrEmpty(p))
            {
                _latestExplorerResult = p;
                return false;
            }
            return true;
        }

        // 线程隔离的回调结果
        [ThreadStatic]
        private static string _latestExplorerResult;

        public static bool IsExplorerClass(string cls)
        {
            if (string.IsNullOrEmpty(cls)) return false;
            return cls == "CabinetWClass"
                || cls == "ExplorerWClass"
                || cls == "Progman"
                || cls == "WorkerW";
        }

        public static string GetClassNameString(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        /// <summary>
        /// 通过 UI Automation 读取 Explorer 地址栏文本。
        /// 必须在 STA 线程上调用。
        /// 实际实现优先走 Shell COM (更稳定), 失败时回落 UI Automation.
        /// </summary>
        public static string GetPathFromExplorerWindow(IntPtr hwnd)
        {
            try
            {
                string shellPath = GetExplorerPathViaShell(hwnd);
                if (!string.IsNullOrEmpty(shellPath) && IsValidPath(shellPath))
                {
                    return shellPath;
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("GetPathFromExplorerWindow Shell EX: " + ex.Message);
            }

            try
            {
                string msaaPath = GetExplorerPathViaAccessible(hwnd);
                if (!string.IsNullOrEmpty(msaaPath) && IsValidPath(msaaPath))
                {
                    return msaaPath;
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("GetPathFromExplorerWindow MSAA EX: " + ex.Message);
            }

            try
            {
                AutomationElement element = AutomationElement.FromHandle(hwnd);
                if (element == null) return null;

                AutomationElementCollection edits;
                try
                {
                    edits = element.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                }
                catch
                {
                    return null;
                }

                foreach (AutomationElement edit in edits)
                {
                    try
                    {
                        if (edit.Current.Name != null && IsValidPath(edit.Current.Name))
                        {
                            return edit.Current.Name;
                        }

                        ValuePattern vp = edit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        if (vp != null && !string.IsNullOrEmpty(vp.Current.Value) && IsValidPath(vp.Current.Value))
                        {
                            return vp.Current.Value;
                        }
                    }
                    catch
                    {
                        // skip element
                    }
                }
            }
            catch
            {
                // UI Automation 在某些 explorer 状态下会抛异常, 这里吃掉
            }
            return null;
        }

        public static bool IsValidPath(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            if (s.Length < 3) return false;
            if (s.Contains("\\") && !s.Contains("?") && !s.Contains("*")) return true;
            return false;
        }

        /// <summary>
        /// 判断是否应当跳过此对话框 (自身进程 / 缓存中已存在)。
        /// </summary>
        public static bool ShouldSkipDialog(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return true;
            if (!IsWindow(hwnd)) return true;

            try
            {
                GetWindowThreadProcessId(hwnd, out uint pid);
                if (pid == _geekDeskPid) return true;
            }
            catch
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 在 #32770 对话框中查找地址栏 Edit 控件并尝试注入路径。
        /// </summary>
        public static bool TryInject(IntPtr dialogHwnd, string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                LogUtil.WriteQuickSwitchLog("TryInject: empty path");
                return false;
            }
            if (!IsWindow(dialogHwnd))
            {
                LogUtil.WriteQuickSwitchLog("TryInject: dialog hwnd invalid");
                return false;
            }

            IntPtr editHwnd = FindAddressEdit(dialogHwnd);
            if (editHwnd == IntPtr.Zero)
            {
                LogUtil.WriteQuickSwitchLog("TryInject: no Edit found in dialog 0x" + dialogHwnd.ToString("X"));
                return false;
            }

            string current = GetEditText(editHwnd);
            if (string.Equals(current, path, StringComparison.OrdinalIgnoreCase))
            {
                LogUtil.WriteQuickSwitchLog("TryInject: path already equals current, skip. hwnd=0x" + editHwnd.ToString("X") + " path=" + path);
                return false;
            }

            try
            {
                IntPtr setTextResult = SendMessage(editHwnd, WM_SETTEXT, IntPtr.Zero, path);
                bool postDown = PostMessage(editHwnd, WM_KEYDOWN, (IntPtr)VK_RETURN, IntPtr.Zero);
                bool postUp = PostMessage(editHwnd, WM_KEYUP, (IntPtr)VK_RETURN, IntPtr.Zero);
                LogUtil.WriteQuickSwitchLog("TryInject: OK dialog=0x" + dialogHwnd.ToString("X") + " edit=0x" + editHwnd.ToString("X")
                    + " SetTextRet=" + setTextResult.ToInt64()
                    + " PostDown=" + postDown + " PostUp=" + postUp
                    + " path=" + path);
                return true;
            }
            catch (Exception ex)
            {
                LogUtil.WriteErrorLog(ex, "QuickSwitch 注入失败");
                return false;
            }
        }

        public static string GetEditText(IntPtr editHwnd)
        {
            try
            {
                var sb = new StringBuilder(1024);
                GetWindowText(editHwnd, sb, sb.Capacity);
                return sb.ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 在 #32770 子窗口树中查找地址栏 Edit。
        /// 优先按"文本内容是路径"的 Edit 匹配 (兼容 Win10 老 / Win11 新两种对话框),
        /// 然后退化到 ComboBoxEx32 → ComboBox → Edit 链, 最后 fallback 到第一个 Edit。
        /// 必须在 STA 线程上调用 (UI Automation 需要 STA COM apartment)。
        /// </summary>
        public static IntPtr FindAddressEdit(IntPtr dialogHwnd)
        {
            var candidates = new System.Collections.Generic.List<KeyValuePair<IntPtr, string>>();

            try
            {
                AutomationElement element = AutomationElement.FromHandle(dialogHwnd);
                if (element != null)
                {
                    AutomationElementCollection edits = element.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

                    foreach (AutomationElement edit in edits)
                    {
                        try
                        {
                            IntPtr h = new IntPtr(edit.Current.NativeWindowHandle);
                            if (h == IntPtr.Zero || !IsWindow(h)) continue;

                            string text = null;
                            try { text = edit.Current.Name; } catch { }
                            if (string.IsNullOrEmpty(text))
                            {
                                try
                                {
                                    ValuePattern vp = edit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                    if (vp != null) text = vp.Current.Value;
                                }
                                catch { }
                            }

                            candidates.Add(new KeyValuePair<IntPtr, string>(h, text ?? ""));

                            // 命中: Edit 文本看起来像路径 (含 drive letter + "\\")
                            if (!string.IsNullOrEmpty(text) && text.Length >= 3
                                && text[1] == ':' && text.Contains("\\") && !text.Contains("?") && !text.Contains("*"))
                            {
                                LogUtil.WriteQuickSwitchLog("FindAddressEdit: matched by text '" + text + "' hwnd=0x" + h.ToString("X"));
                                return h;
                            }
                        }
                        catch
                        {
                            // ignored
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("FindAddressEdit UI Automation EX: " + ex.Message);
            }

            // 退化: 遍历 ComboBoxEx32 -> ComboBox -> Edit 链
            IntPtr comboEx = FindFirstChildByClass(dialogHwnd, "ComboBoxEx32");
            if (comboEx != IntPtr.Zero)
            {
                IntPtr combo = FindFirstChildByClass(comboEx, "ComboBox");
                if (combo != IntPtr.Zero)
                {
                    IntPtr edit = FindFirstChildByClass(combo, "Edit");
                    if (edit != IntPtr.Zero)
                    {
                        LogUtil.WriteQuickSwitchLog("FindAddressEdit: matched by ComboBoxEx32 chain hwnd=0x" + edit.ToString("X"));
                        return edit;
                    }
                }
            }

            // 再退化: 顶层 ComboBox -> Edit (Win11 新对话框)
            IntPtr topCombo = FindFirstChildByClass(dialogHwnd, "ComboBox");
            if (topCombo != IntPtr.Zero)
            {
                IntPtr edit = FindFirstChildByClass(topCombo, "Edit");
                if (edit != IntPtr.Zero)
                {
                    LogUtil.WriteQuickSwitchLog("FindAddressEdit: matched by ComboBox->Edit chain hwnd=0x" + edit.ToString("X"));
                    return edit;
                }
            }

            // 最后一搏: 取第一个 Edit (可能错给文件名框, 但起码不会注入失败)
            foreach (var kv in candidates)
            {
                if (kv.Key != IntPtr.Zero)
                {
                    LogUtil.WriteQuickSwitchLog("FindAddressEdit: fallback first Edit text='" + kv.Value + "' hwnd=0x" + kv.Key.ToString("X"));
                    return kv.Key;
                }
            }

            LogUtil.WriteQuickSwitchLog("FindAddressEdit: NO Edit found in dialog 0x" + dialogHwnd.ToString("X") + " (candidates=" + candidates.Count + ")");
            return IntPtr.Zero;
        }

        private static IntPtr FindFirstChildByClass(IntPtr parent, string className)
        {
            IntPtr result = IntPtr.Zero;
            try
            {
                EnumChildWindows(parent, (hwnd, lParam) =>
                {
                    if (result != IntPtr.Zero) return false;
                    string cls = GetClassNameString(hwnd);
                    if (cls == className)
                    {
                        result = hwnd;
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);
            }
            catch { }
            return result;
        }
    }
}
