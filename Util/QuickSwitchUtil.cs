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

        #region MSAA (IAccessible) 备用读取 Explorer 地址栏

        /// <summary>
        /// 通过 MSAA / IAccessible 读取 Explorer 窗口地址栏。
        /// 对 Win10/Win11 都稳定 (UI Automation 在某些场景会失败, MSAA 仍可用)。
        /// 必须 STA 线程。
        /// </summary>
        public static string GetExplorerPathViaAccessible(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd)) return null;

            IntPtr pAcc = IntPtr.Zero;
            try
            {
                Guid iid = IID_IAccessible;
                int hr = AccessibleObjectFromWindow(hwnd, OBJID_WINDOW, ref iid, out pAcc);
                if (hr != 0 || pAcc == IntPtr.Zero) return null;

                // 用反射调 IAccessible.get_accValue / get_accName (避免引用 Interop.UIAutomationClient 等)
                var acc = Marshal.GetObjectForIUnknown(pAcc);
                if (acc == null) return null;

                Type t = acc.GetType();
                string found = null;
                found = WalkAccessible(acc, t, depth: 0);
                return found;
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("GetExplorerPathViaAccessible EX: " + ex.Message);
                return null;
            }
            finally
            {
                if (pAcc != IntPtr.Zero) Marshal.Release(pAcc);
            }
        }

        private static string WalkAccessible(object acc, Type t, int depth)
        {
            if (acc == null || depth > 6) return null;

            // 读当前 accValue / accName
            try
            {
                object v = t.InvokeMember("accValue", System.Reflection.BindingFlags.GetProperty, null, acc, new object[] { 0 });
                if (v is string s && !string.IsNullOrEmpty(s))
                {
                    string dec = UnescapeMsaa(s);
                    if (IsValidPath(dec)) return dec;
                }
            }
            catch { }
            try
            {
                object n = t.InvokeMember("accName", System.Reflection.BindingFlags.GetProperty, null, acc, new object[] { 0 });
                if (n is string s2 && !string.IsNullOrEmpty(s2))
                {
                    string dec = UnescapeMsaa(s2);
                    if (IsValidPath(dec)) return dec;
                }
            }
            catch { }

            // 遍历子对象 (accChildCount + get_accChild)
            try
            {
                object cc = t.InvokeMember("accChildCount", System.Reflection.BindingFlags.GetProperty, null, acc, null);
                if (cc is int cnt && cnt > 0)
                {
                    for (int i = 1; i <= Math.Min(cnt, 200); i++)
                    {
                        object child = null;
                        try
                        {
                            child = t.InvokeMember("get_accChild", System.Reflection.BindingFlags.InvokeMethod, null, acc, new object[] { i });
                        }
                        catch { }
                        if (child == null) continue;

                        // child 本身就是 IAccessible 代理, 直接 WalkAccessible
                        if (child == null) continue;
                        string r = WalkAccessible(child, child.GetType(), depth + 1);
                        if (!string.IsNullOrEmpty(r)) return r;
                    }
                }
            }
            catch { }

            return null;
        }

        private static string UnescapeMsaa(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            try { s = Uri.UnescapeDataString(s); } catch { }
            return s.Replace('/', '\\');
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
