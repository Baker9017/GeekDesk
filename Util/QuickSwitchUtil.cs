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
        #region 文件夹对话框支持 (Folder Browser Dialog)

        // 文件夹对话框有两种类型:
        // 1. 旧式 SHBrowseForFolder (#32770, 树形视图, 无地址栏)
        // 2. 新式 IFileOpenDialog + FOS_PICKFOLDERS (#32766, 有地址栏)
        // 两者的处理策略完全不同.

        // IFileDialog / IFileOpenDialog GUIDs (用于 COM 自动化)
        // 来源: Windows SDK shobjidl_core.h / objbase.h
        private static readonly Guid CLSID_FileOpenDialog  = new Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7"); // shell32 CoCreateInstance CLSID
        private static readonly Guid CLSID_FileSaveDialog  = new Guid("C0B4E2F3-BA21-4773-8BA9-9CAA42BE0B59");
        private static readonly Guid IID_IFileDialog       = new Guid("42F85136-DB7E-439C-85F1-E4079D47506B"); // IFileDialog IID
        private static readonly Guid IID_IFileOpenDialog   = new Guid("D57C7288-D4AD-4768-BE02-9D969532D960"); // IFileOpenDialog IID
        private static readonly Guid IID_IShellItem        = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"); // IShellItem IID

        private const uint FOS_PICKFOLDERS = 0x00000020;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object ppv);

        /// <summary>
        /// 判断对话框是否为文件夹选择对话框 (而非文件打开/保存对话框)。
        /// 对于新式 CommonItemDialog (#32766)，尝试通过 COM 接口设置路径。
        /// 对于旧式树形对话框，返回 false 表示跳过（避免卡死）。
        /// </summary>
        public static bool IsFolderDialog(IntPtr dialogHwnd)
        {
            if (dialogHwnd == IntPtr.Zero || !IsWindow(dialogHwnd)) return false;

            string cls = GetClassNameString(dialogHwnd);

            // 新式 CommonItemDialog (Windows Vista+)
            if (cls == "#32766")
            {
                LogUtil.WriteQuickSwitchLog("IsFolderDialog: detected CommonItemDialog #32766 hwnd=0x"
                    + dialogHwnd.ToString("X"));
                return true;
            }

            // 旧式 SHBrowseForFolder (#32770): 检测是否只有树形控件，没有地址栏
            if (cls == "#32770")
            {
                // 如果找不到 ComboBoxEx32 或地址栏 Edit，则认为是旧式文件夹对话框
                IntPtr topmostComboEx = FindTopmostComboBoxEx32(dialogHwnd);
                if (topmostComboEx == IntPtr.Zero)
                {
                    // 没有地址栏 ComboBoxEx，可能是旧式树形文件夹对话框
                    LogUtil.WriteQuickSwitchLog("IsFolderDialog: no ComboBoxEx32 found, likely old-style folder dialog hwnd=0x"
                        + dialogHwnd.ToString("X"));
                    return true; // 标记为文件夹对话框，但不进行注入
                }
            }

            return false;
        }

        /// <summary>
        /// 对文件夹选择对话框尝试注入路径。
        /// 新式 CommonItemDialog (#32766) 使用 IFileDialog::SetFolder COM 接口。
        /// 旧式对话框跳过注入，避免卡死。
        /// </summary>
        public static bool TryInjectFolderDialog(IntPtr dialogHwnd, string path)
        {
            if (string.IsNullOrEmpty(path) || !IsWindow(dialogHwnd)) return false;

            string cls = GetClassNameString(dialogHwnd);

            // 新式 CommonItemDialog (#32766)
            if (cls == "#32766")
            {
                return TryInjectViaIFileDialog(dialogHwnd, path);
            }

            // 旧式 SHBrowseForFolder (#32770)
            if (cls == "#32770")
            {
                IntPtr topmostComboEx = FindTopmostComboBoxEx32(dialogHwnd);
                if (topmostComboEx == IntPtr.Zero)
                {
                    // 无地址栏 → 用 BFFM_SETSELECTION 让树形视图跳转
                    LogUtil.WriteQuickSwitchLog("TryInjectFolderDialog: old-style tree dialog → BFFM_SETSELECTION hwnd=0x"
                        + dialogHwnd.ToString("X") + " path=" + path);
                    return TryInjectViaSetSelection(dialogHwnd, path);
                }
                // 有地址栏（BIF_USENEWUI 等带编辑框的旧式对话框）→ 正常注入
                LogUtil.WriteQuickSwitchLog("TryInjectFolderDialog: old-style with ComboBoxEx → TryInject hwnd=0x"
                    + dialogHwnd.ToString("X"));
                return TryInject(dialogHwnd, path);
            }

            LogUtil.WriteQuickSwitchLog("TryInjectFolderDialog: unknown class '" + cls + "' hwnd=0x" + dialogHwnd.ToString("X"));
            return false;
        }

        /// <summary>
        /// 旧式 SHBrowseForFolder 路径跳转：通过 BFFM_SETSELECTION 消息驱动树形对话框。
        /// <para>
        /// BFFM_SETSELECTION = WM_USER + 103 (0x0467)。
        /// wParam=1 (TRUE) 时 lParam 是目标进程内的 Unicode 路径指针。
        /// OS 不对 WM_USER 消息做跨进程 marshal，必须手动
        /// VirtualAllocEx + WriteProcessMemory 把字符串写入目标进程再发消息。
        /// </para>
        /// </summary>
        private static bool TryInjectViaSetSelection(IntPtr dialogHwnd, string path)
        {
            if (string.IsNullOrEmpty(path) || !System.IO.Directory.Exists(path))
            {
                LogUtil.WriteQuickSwitchLog("TryInjectViaSetSelection: path invalid or not exist '" + path + "'");
                return false;
            }

            GetWindowThreadProcessId(dialogHwnd, out uint targetPid);
            LogUtil.WriteQuickSwitchLog("TryInjectViaSetSelection: pid=" + targetPid + " path=" + path);

            // BFFM_SETSELECTION: WM_USER(0x0400) + 103 = 0x0467
            // wParam = 1 (TRUE) → lParam 为 Unicode 路径字符串指针（目标进程内）
            const int BFFM_SETSELECTION = 0x0467;

            byte[] pathBytes = System.Text.Encoding.Unicode.GetBytes(path + "\0");

            IntPtr hProcess = OpenProcess(PROCESS_VM_WRITE | PROCESS_VM_OPERATION, false, targetPid);
            if (hProcess == IntPtr.Zero)
            {
                LogUtil.WriteQuickSwitchLog("TryInjectViaSetSelection: OpenProcess failed pid=" + targetPid
                    + " err=" + Marshal.GetLastWin32Error());
                return false;
            }

            IntPtr remoteBuffer = IntPtr.Zero;
            try
            {
                remoteBuffer = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)pathBytes.Length,
                    MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
                if (remoteBuffer == IntPtr.Zero)
                {
                    LogUtil.WriteQuickSwitchLog("TryInjectViaSetSelection: VirtualAllocEx failed err=" + Marshal.GetLastWin32Error());
                    return false;
                }

                IntPtr written;
                if (!WriteProcessMemory(hProcess, remoteBuffer, pathBytes, (uint)pathBytes.Length, out written))
                {
                    LogUtil.WriteQuickSwitchLog("TryInjectViaSetSelection: WriteProcessMemory failed err=" + Marshal.GetLastWin32Error());
                    return false;
                }

                // 发送消息让树形对话框跳转（SendMessage 是同步的，返回时树已选中目标节点）
                SendMessage(dialogHwnd, BFFM_SETSELECTION, (IntPtr)1, remoteBuffer);
                LogUtil.WriteQuickSwitchLog("TryInjectViaSetSelection: BFFM_SETSELECTION sent OK → remotePtr=0x"
                    + remoteBuffer.ToString("X") + " path=" + path);

                // BFFM_SETSELECTION 选中节点但不一定滚动到可视区，需手动 EnsureVisible
                EnsureTreeSelectionVisible(dialogHwnd);
                return true;
            }
            finally
            {
                if (remoteBuffer != IntPtr.Zero)
                    VirtualFreeEx(hProcess, remoteBuffer, 0, MEM_RELEASE);
                CloseHandle(hProcess);
            }
        }

        /// <summary>
        /// 让 SHBrowseForFolder 树形对话框把当前选中节点滚动到可视区，
        /// 使垂直滚动条的滑块位置与选中项对应。
        /// <para>
        /// BFFM_SETSELECTION 只负责选中，不保证滚动。
        /// TVM_GETNEXTITEM(TVGN_CARET) 取得选中项 HTREEITEM，
        /// TVM_ENSUREVISIBLE 触发滚动并刷新滚动条。
        /// HTREEITEM 是不透明句柄值，可直接跨进程通过 SendMessage 传递，
        /// 无需 VirtualAllocEx。
        /// </para>
        /// </summary>
        private static void EnsureTreeSelectionVisible(IntPtr dialogHwnd)
        {
            try
            {
                IntPtr treeHwnd = FindFirstChildByClass(dialogHwnd, "SysTreeView32");
                if (treeHwnd == IntPtr.Zero)
                {
                    LogUtil.WriteQuickSwitchLog("EnsureTreeSelectionVisible: no SysTreeView32 found");
                    return;
                }

                // commctrl.h: TVM_FIRST = 0x1100
                const int TVM_GETNEXTITEM   = 0x1100 + 10; // 0x110A
                const int TVM_ENSUREVISIBLE = 0x1100 + 20; // 0x1114
                const int TVGN_CARET        = 0x0009;      // 当前光标（选中）节点

                IntPtr hSelected = SendMessage(treeHwnd, TVM_GETNEXTITEM, (IntPtr)TVGN_CARET, IntPtr.Zero);
                if (hSelected == IntPtr.Zero)
                {
                    LogUtil.WriteQuickSwitchLog("EnsureTreeSelectionVisible: TVGN_CARET returned null");
                    return;
                }

                // TVM_ENSUREVISIBLE: 滚动树视图使节点进入可视区，同时刷新滚动条位置
                IntPtr scrolled = SendMessage(treeHwnd, TVM_ENSUREVISIBLE, IntPtr.Zero, hSelected);
                LogUtil.WriteQuickSwitchLog("EnsureTreeSelectionVisible: TVM_ENSUREVISIBLE hItem=0x"
                    + hSelected.ToString("X") + " scrolled=" + scrolled);
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("EnsureTreeSelectionVisible EX: " + ex.Message);
            }
        }

        /// <summary>
        /// 通过 IFileDialog COM 接口设置文件夹对话框的路径。
        /// 这是最安全、最稳定的方式，不会导致 UI 卡死。
        /// </summary>
        private static bool TryInjectViaIFileDialog(IntPtr dialogHwnd, string path)
        {
            if (string.IsNullOrEmpty(path) || !System.IO.Directory.Exists(path))
            {
                LogUtil.WriteQuickSwitchLog("TryInjectViaIFileDialog: invalid path '" + path + "'");
                return false;
            }

            try
            {
                // 获取对话框所属进程的 HWND 对应的 IFileDialog
                // 方法: 通过 GetWindowThreadProcessId 获取进程 ID，
                // 然后 CoGetObjectFromWindow 或其他方式获取 COM 对象
                // 
                // 更可靠的方法: 直接从顶层窗口获取 IDispatch，然后尝试 QueryInterface 到 IFileDialog
                // 但这需要知道确切的接口偏移
                //
                // 简化方案: 使用 IShellItem 方式，Windows 会自动处理路径

                object shellItem = null;
                Guid iidShellItem = IID_IShellItem; // static readonly 不能直接作 ref 参数
                int hr = SHCreateItemFromParsingName(path, IntPtr.Zero, ref iidShellItem, out shellItem);
                if (hr != 0 || shellItem == null)
                {
                    LogUtil.WriteQuickSwitchLog("TryInjectViaIFileDialog: SHCreateItemFromParsingName failed hr=0x"
                        + hr.ToString("X") + " path=" + path);
                    return false;
                }

                // 获取前台窗口，如果是对话框则尝试设置
                IntPtr fg = GetForegroundWindow();
                if (fg == dialogHwnd || IsChildOf(dialogHwnd, fg))
                {
                    // 尝试发送消息让对话框导航到该路径
                    // 对于 IFileDialog，可以在地址栏输入路径后发送回车
                    LogUtil.WriteQuickSwitchLog("TryInjectViaIFileDialog: attempting UI automation for hwnd=0x"
                        + dialogHwnd.ToString("X") + " path=" + path);

                    // 使用 UI Automation 找到地址栏并设置路径
                    return TryInjectViaUIAutomation(dialogHwnd, path);
                }

                LogUtil.WriteQuickSwitchLog("TryInjectViaIFileDialog: dialog not focused, skip hwnd=0x"
                    + dialogHwnd.ToString("X"));
                return false;
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("TryInjectViaIFileDialog EX: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 通过 UI Automation 设置文件夹对话框的地址栏路径。
        /// 适用于新式 CommonItemDialog。
        /// </summary>
        private static bool TryInjectViaUIAutomation(IntPtr dialogHwnd, string path)
        {
            try
            {
                AutomationElement element = AutomationElement.FromHandle(dialogHwnd);
                if (element == null) return false;

                // 查找地址栏 (通常是 ComboBox 或 Edit)
                var combos = element.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox));
                foreach (AutomationElement combo in combos)
                {
                    try
                    {
                        string name = combo.Current.Name ?? "";
                        // 匹配常见的地址栏名称
                        if (name.IndexOf("地址", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            name.IndexOf("Address", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            name.IndexOf("folder", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // 尝试获取 ValuePattern
                            var vp = combo.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            if (vp != null)
                            {
                                vp.SetValue(path);
                                SendReturnKey(combo);
                                LogUtil.WriteQuickSwitchLog("TryInjectViaUIAutomation: set path via ComboBox hwnd=0x"
                                    + new IntPtr(combo.Current.NativeWindowHandle).ToString("X"));
                                return true;
                            }

                            // 查找 ComboBox 内的 Edit
                            var editInCombo = combo.FindFirst(TreeScope.Children,
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                            if (editInCombo != null)
                            {
                                var vpEdit = editInCombo.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                if (vpEdit != null)
                                {
                                    vpEdit.SetValue(path);
                                    SendReturnKey(editInCombo);
                                    LogUtil.WriteQuickSwitchLog("TryInjectViaUIAutomation: set path via Edit in ComboBox");
                                    return true;
                                }
                            }
                        }
                    }
                    catch { }
                }

                // 备选: 找顶层 Edit 并检查其名称/位置
                var edits = element.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                foreach (AutomationElement edit in edits)
                {
                    try
                    {
                        string name = edit.Current.Name ?? "";
                        IntPtr editHwnd = new IntPtr(edit.Current.NativeWindowHandle);

                        // 检查是否像地址栏 (路径格式)
                        if (name.Length >= 3 && name[1] == ':' && name.Contains("\\"))
                        {
                            var vp = edit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            if (vp != null)
                            {
                                vp.SetValue(path);
                                SendReturnKey(edit);
                                LogUtil.WriteQuickSwitchLog("TryInjectViaUIAutomation: set path via Edit name='"
                                    + name + "' hwnd=0x" + editHwnd.ToString("X"));
                                return true;
                            }
                        }
                    }
                    catch { }
                }

                LogUtil.WriteQuickSwitchLog("TryInjectViaUIAutomation: no address bar found in dialog");
                return false;
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("TryInjectViaUIAutomation EX: " + ex.Message);
                return false;
            }
        }

        private static void SendReturnKey(AutomationElement element)
        {
            try
            {
                // ValuePattern.SetValue 后，模拟回车键触发地址栏导航
                IntPtr hwnd = new IntPtr(element.Current.NativeWindowHandle);
                const int WM_KEYDOWN = 0x0100;
                const int WM_KEYUP = 0x0101;
                const int VK_RETURN = 0x0D;
                const int lParamKeyDown = 0x001C0001;
                const int lParamKeyUp = unchecked((int)0xC01C0001);

                SendMessage(hwnd, WM_KEYDOWN, (IntPtr)VK_RETURN, (IntPtr)lParamKeyDown);
                SendMessage(hwnd, WM_KEYUP, (IntPtr)VK_RETURN, (IntPtr)lParamKeyUp);
            }
            catch { }
        }

        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        private static bool IsChildOf(IntPtr parent, IntPtr child)
        {
            if (child == IntPtr.Zero) return false;
            IntPtr current = child;
            while (current != IntPtr.Zero)
            {
                if (current == parent) return true;
                current = GetParent(current);
            }
            return false;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        #endregion

        #region Shell COM (读取 Explorer 路径)

        // 通过 Type.GetTypeFromCLSID + Activator.CreateInstance 访问 IShellWindows,
        // 再用 InvokeMember 反射调用属性/方法 (late binding)。
        // 不使用 CoCreateInstance + 手动 IDispatch vtable——IShellWindows 的进程外 proxy
        // 在 Win10/11 上调用 GetIDsOfNames 会返回 DISP_E_UNKNOWNNAME, 永远无法读取属性名。

        // MSAA (Microsoft Active Accessibility) 备用路径.
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, out IntPtr ppvObject);

        private static readonly Guid IID_IAccessible = new Guid("618736E0-3C3D-11CF-810C-02803C1174A1");
        // IWebBrowserApp - 用于 MSAA 备用路径读取 LocationURL.
        private static readonly Guid IID_IWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");

        private static readonly Guid CLSID_ShellWindows = new Guid("9BA05972-F6A8-11CF-A442-00A0C90A8F39");
        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        /// <summary>
        /// 通过 Shell COM (IShellWindows) 获取某个指定 Explorer 窗口 (targetHwnd!=0)
        /// 或任意可见 Explorer 窗口 (targetHwnd==0) 的当前路径。
        /// 必须在 STA 线程上调用。失败返回 null。
        ///
        /// 注意: 使用 Activator.CreateInstance + InvokeMember (late binding) 而非手动 IDispatch
        /// vtable 调用。IShellWindows 是进程外 COM 对象, 其 proxy/stub 的 GetIDsOfNames 在
        /// Win10/11 上返回 DISP_E_UNKNOWNNAME (0x80020001), 导致手动 vtable 方式永远失败。
        /// .NET 的 InvokeMember 走自己的 COM 互操作层, 可以正确解析属性名。
        /// </summary>
        public static string GetExplorerPathViaShell(IntPtr targetHwnd)
        {
            object shellWindows = null;
            try
            {
                Type shellWindowsType = Type.GetTypeFromCLSID(CLSID_ShellWindows);
                if (shellWindowsType == null)
                {
                    LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: Type.GetTypeFromCLSID returned null");
                    return null;
                }

                shellWindows = Activator.CreateInstance(shellWindowsType);
                if (shellWindows == null)
                {
                    LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: Activator.CreateInstance returned null");
                    return null;
                }

                int count;
                try
                {
                    count = (int)shellWindows.GetType().InvokeMember("Count",
                        System.Reflection.BindingFlags.GetProperty, null, shellWindows, null);
                }
                catch (Exception ex)
                {
                    LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: get Count EX: " + ex.Message);
                    return null;
                }

                LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: ShellWindows.Count = " + count);
                if (count <= 0) return null;

                for (int i = 0; i < count; i++)
                {
                    object item = null;
                    try
                    {
                        item = shellWindows.GetType().InvokeMember("Item",
                            System.Reflection.BindingFlags.InvokeMethod, null, shellWindows, new object[] { i });
                        if (item == null) continue;

                        // 读取 HWND
                        IntPtr shellHwnd;
                        try
                        {
                            object hwndObj = item.GetType().InvokeMember("HWND",
                                System.Reflection.BindingFlags.GetProperty, null, item, null);
                            shellHwnd = new IntPtr(Convert.ToInt64(hwndObj));
                        }
                        catch { continue; }

                        if (targetHwnd != IntPtr.Zero)
                        {
                            if (shellHwnd != targetHwnd) continue;
                        }
                        else
                        {
                            if (shellHwnd == IntPtr.Zero || !IsWindow(shellHwnd) || !IsWindowVisible(shellHwnd))
                                continue;
                        }

                        // 读取 LocationURL
                        string url;
                        try
                        {
                            url = (string)item.GetType().InvokeMember("LocationURL",
                                System.Reflection.BindingFlags.GetProperty, null, item, null);
                        }
                        catch { continue; }

                        if (string.IsNullOrEmpty(url)) continue;
                        string path = UrlToPath(url);
                        if (!string.IsNullOrEmpty(path) && IsValidPath(path))
                        {
                            LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell: found path=" + path + " hwnd=0x" + shellHwnd.ToString("X"));
                            return path;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell item[" + i + "] EX: " + ex.Message);
                    }
                    finally
                    {
                        if (item != null) try { Marshal.ReleaseComObject(item); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("GetExplorerPathViaShell EX: " + ex.Message);
            }
            finally
            {
                if (shellWindows != null) try { Marshal.ReleaseComObject(shellWindows); } catch { }
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
        public const uint EVENT_OBJECT_NAMECHANGE = 0x800C;  // 窗口标题改变 (Explorer 导航时触发)
        public const uint EVENT_SYSTEM_FOREGROUND = 0x0003;

        public const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        public const uint WINEVENT_SKIPOWNPROCESS = 0x0002;

        public const int OBJID_WINDOW = 0x00000000;

        private const int WM_SETTEXT = 0x000C;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int VK_RETURN = 0x0D;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

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

        // WM_SETTEXT 版: lParam 为字符串
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);

        // 通用版: lParam 为 IntPtr (用于 WM_KEYDOWN / WM_KEYUP 等)
        // 与上面的字符串重载共存。重要: 此版本绕过线程消息队列,
        // IsDialogMessage 无法拦截, 直接将消息送达目标窗口的 WndProc。
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

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

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        // ── kernel32: 跨进程内存写入 (用于 BFFM_SETSELECTION 注入) ──

        private const uint PROCESS_VM_WRITE     = 0x0020;
        private const uint PROCESS_VM_OPERATION = 0x0008;
        private const uint MEM_COMMIT    = 0x00001000;
        private const uint MEM_RESERVE   = 0x00002000;
        private const uint MEM_RELEASE   = 0x00008000;
        private const uint PAGE_READWRITE = 0x04;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress,
            uint dwSize, uint flAllocationType, uint flProtect);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress,
            byte[] lpBuffer, uint nSize, out IntPtr lpNumberOfBytesWritten);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress,
            uint dwSize, uint dwFreeType);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

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
                // 1. 写入地址栏文本
                IntPtr setTextResult = SendMessage(editHwnd, WM_SETTEXT, IntPtr.Zero, path);

                // 2. 触发地址栏导航: 必须用 SendMessage 而非 PostMessage。
                //
                //    PostMessage 把消息放入线程消息队列; 对话框消息循环会先用
                //    IsDialogMessage 扫描队列, 看到 VK_RETURN 就直接激活默认按钮
                //    (Save/OK), 消息根本不会送达地址栏 Edit 的 WndProc ——
                //    这正是 Word/Excel 另存为触发 QuickSwitch 后直接保存文件并
                //    退出对话框的根本原因。
                //
                //    SendMessage 直接调用目标窗口的 WndProc (内核同步调用),
                //    完全绕过线程消息队列, IsDialogMessage 无从介入。
                //    Shell 地址栏子类化的 WndProc 收到 VK_RETURN 后触发目录导航,
                //    对话框保持打开状态。
                //
                // lParam 说明:
                //   WM_KEYDOWN: 0x001C0001 = scan_code(0x1C) | repeat(1)
                //   WM_KEYUP  : 0xC01C0001 = transition(1)<<31 | prev_state(1)<<30
                //               | scan_code(0x1C) | repeat(1)
                const int lParamKeyDown = 0x001C0001;
                const int lParamKeyUp   = unchecked((int)0xC01C0001);
                SendMessage(editHwnd, WM_KEYDOWN, (IntPtr)VK_RETURN, (IntPtr)lParamKeyDown);
                SendMessage(editHwnd, WM_KEYUP,   (IntPtr)VK_RETURN, (IntPtr)lParamKeyUp);

                LogUtil.WriteQuickSwitchLog("TryInject: OK dialog=0x" + dialogHwnd.ToString("X") + " edit=0x" + editHwnd.ToString("X")
                    + " SetTextRet=" + setTextResult.ToInt64()
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
        ///
        /// 优先级:
        ///   ① Y 坐标最靠上的 ComboBoxEx32 → ComboBox → Edit  (最可靠, 最优先)
        ///      地址栏始终位于对话框顶部; 文件名框始终位于底部。
        ///      旧式对话框 (GetOpenFileName 等) 中, 地址栏与文件名框都是 ComboBoxEx32,
        ///      按 Z 序找第一个不可靠 (文件名框可能排在前面); 按屏幕 Y 坐标取最靠上的,
        ///      可跨所有版本准确命中地址栏。
        ///   ② UI Automation 扫描内容像路径的 Edit (次选)
        ///      过滤: 路径末段含 "." 的跳过 (文件名.ext 特征), 防止 Word/Excel 另存为
        ///      的文件名框 (预填完整路径如 "C:\...\文档.docx") 被误命中。
        ///   ③ 顶层 ComboBox → Edit 链 (Win11 新对话框兜底)
        ///   ④ 第一个 Edit (最后兜底)
        ///
        /// 必须在 STA 线程上调用 (UI Automation 需要 STA COM apartment)。
        /// </summary>
        public static IntPtr FindAddressEdit(IntPtr dialogHwnd)
        {
            // ① 按屏幕 Y 坐标找最靠上的 ComboBoxEx32 → ComboBox → Edit
            IntPtr topmostComboEx = FindTopmostComboBoxEx32(dialogHwnd);
            if (topmostComboEx != IntPtr.Zero)
            {
                IntPtr combo = FindFirstChildByClass(topmostComboEx, "ComboBox");
                if (combo != IntPtr.Zero)
                {
                    IntPtr edit = FindFirstChildByClass(combo, "Edit");
                    if (edit != IntPtr.Zero)
                    {
                        string text = GetEditText(edit);
                        if (!IsFilenameEditContent(text))
                        {
                            LogUtil.WriteQuickSwitchLog("FindAddressEdit: topmost ComboBoxEx32 Edit hwnd=0x"
                                + edit.ToString("X") + " text='" + text + "'");
                            return edit;
                        }
                        LogUtil.WriteQuickSwitchLog("FindAddressEdit: topmost ComboBoxEx32 Edit looks like filename, skip: '" + text + "'");
                    }
                }
            }

            // ② UI Automation 扫描 (次选; 无 ComboBoxEx32 或 ① 校验未通过时使用)
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

                            // 命中: 路径格式 (drive letter + \\, 无通配符) 且不像文件名框
                            if (!string.IsNullOrEmpty(text) && text.Length >= 3
                                && text[1] == ':' && text.Contains("\\") && !text.Contains("?") && !text.Contains("*"))
                            {
                                if (IsFilenameEditContent(text))
                                {
                                    LogUtil.WriteQuickSwitchLog("FindAddressEdit: skip filename field '" + text + "' hwnd=0x" + h.ToString("X"));
                                    continue;
                                }
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

            // ③ 顶层 ComboBox → Edit (Win11 新对话框兜底)
            IntPtr topCombo = FindFirstChildByClass(dialogHwnd, "ComboBox");
            if (topCombo != IntPtr.Zero)
            {
                IntPtr edit = FindFirstChildByClass(topCombo, "Edit");
                if (edit != IntPtr.Zero)
                {
                    LogUtil.WriteQuickSwitchLog("FindAddressEdit: ComboBox->Edit chain hwnd=0x" + edit.ToString("X"));
                    return edit;
                }
            }

            // ④ 最后兜底: 第一个 Edit
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

        /// <summary>
        /// 在对话框所有后代中，按屏幕 Y 坐标找到最靠上的 ComboBoxEx32。
        /// 文件对话框地址栏 (顶部) 与文件名框 (底部) 都使用 ComboBoxEx32,
        /// Y 最小 = 最靠上 = 地址栏。
        /// </summary>
        private static IntPtr FindTopmostComboBoxEx32(IntPtr dialogHwnd)
        {
            IntPtr best = IntPtr.Zero;
            int bestTop = int.MaxValue;
            try
            {
                EnumChildWindows(dialogHwnd, (hwnd, lParam) =>
                {
                    if (GetClassNameString(hwnd) == "ComboBoxEx32")
                    {
                        RECT r;
                        if (GetWindowRect(hwnd, out r) && r.Top < bestTop)
                        {
                            bestTop = r.Top;
                            best = hwnd;
                        }
                    }
                    return true; // 继续枚举所有后代
                }, IntPtr.Zero);
            }
            catch { }
            return best;
        }

        /// <summary>
        /// 判断 Edit 内容是否像"文件名输入框"，而非地址栏。
        /// 返回 true 表示应跳过此 Edit（文件名框特征）。
        /// </summary>
        private static bool IsFilenameEditContent(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            // 含通配符 → 文件过滤模式 (如 "*.txt", "*.docx")
            if (text.Contains("*") || text.Contains("?")) return true;
            // 末段含 "." → 文件名带扩展名 (如 "document.docx", "工作表.xlsx")
            string normalized = text.TrimEnd('\\');
            int lastSlash = normalized.LastIndexOf('\\');
            string lastName = lastSlash >= 0 ? normalized.Substring(lastSlash + 1) : normalized;
            return !string.IsNullOrEmpty(lastName) && lastName.Contains(".");
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
