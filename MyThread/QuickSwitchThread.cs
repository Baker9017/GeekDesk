using GeekDesk.Util;
using System;
using System.Windows.Threading;

namespace GeekDesk.MyThread
{
    /// <summary>
    /// Quick Switch (文件夹跟随) 后台线程编排。
    ///
    /// 设计原则
    /// ─────────
    /// 1. 惰性监控: Explorer WinEvent Hook 仅当有 #32770 文件对话框打开时才启动,
    ///    最后一个对话框关闭后立即停止, 不占用任何全局事件资源.
    ///
    /// 2. 延迟注入: Explorer 导航事件只把新路径记录为"待注入" (pending), 不立即
    ///    修改对话框. 当用户把焦点切回文件对话框 (EVENT_SYSTEM_FOREGROUND) 时才注入.
    ///
    /// 3. 首次立即注入: 对话框刚打开时, 仍然立即注入当前 Explorer 路径, 保持跟随体验.
    ///
    /// 4. 无轮询: 去掉了定时轮询线程, 完全由事件驱动.
    /// </summary>
    public static class QuickSwitchThread
    {
        // 始终开启: 监听 #32770 对话框创建/销毁
        private static IntPtr _hookDialog;

        // 惰性开启: 仅当有对话框打开时才存在
        private static IntPtr _hookExplorer;

        private static Dispatcher _staDispatcher;

        private static readonly QuickSwitchUtil.WinEventDelegate _dialogProc       = DialogLifecycleCallback;
        private static readonly QuickSwitchUtil.WinEventDelegate _explorerFocusProc = ExplorerAndFocusCallback;

        /// <summary>
        /// 对话框 HWND → 待注入路径.
        /// 有 key 且 value 非空 = 有待注入的路径变更; key 不存在或 value 为 null = 无待注入.
        /// </summary>
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<IntPtr, string> _pendingPaths
            = new System.Collections.Concurrent.ConcurrentDictionary<IntPtr, string>();

        private static readonly object _lock = new object();
        private static bool _running = false;

        private static System.Windows.Threading.DispatcherTimer _pumpTimer;

        // ─────────────────────────── 生命周期 ───────────────────────────

        public static void Hook()
        {
            lock (_lock)
            {
                if (_running) return;
                _running = true;
            }

            try
            {
                QuickSwitchUtil.InitSelfPid();
                _staDispatcher = DispatcherBuild.Build();

                // 只监听 #32770 对话框生命周期 (始终开启, 开销极低)
                _hookDialog = QuickSwitchUtil.SetWinEventHook(
                    QuickSwitchUtil.EVENT_OBJECT_CREATE,
                    QuickSwitchUtil.EVENT_OBJECT_DESTROY,
                    IntPtr.Zero,
                    _dialogProc,
                    0, 0,
                    QuickSwitchUtil.WINEVENT_OUTOFCONTEXT | QuickSwitchUtil.WINEVENT_SKIPOWNPROCESS);

                // Pump: 在 STA dispatcher 上周期性处理对话框事件队列
                _pumpTimer = new System.Windows.Threading.DispatcherTimer(
                    System.Windows.Threading.DispatcherPriority.Background,
                    _staDispatcher);
                _pumpTimer.Interval = TimeSpan.FromMilliseconds(200);
                _pumpTimer.Tick += (s, e) => ProcessDialogEvents();
                _pumpTimer.Start();

                LogUtil.WriteQuickSwitchLog("Hook 启动: selfPid=" + QuickSwitchUtil.GeekDeskPid
                    + " hookDialog=" + _hookDialog
                    + " pumpThread=" + (_staDispatcher?.Thread?.ManagedThreadId.ToString() ?? "null"));
            }
            catch (Exception ex)
            {
                LogUtil.WriteErrorLog(ex, "QuickSwitch Hook 失败");
                Dispose();
            }
        }

        public static void Dispose()
        {
            lock (_lock)
            {
                if (!_running) return;
                _running = false;
            }

            try { if (_hookDialog   != IntPtr.Zero) { QuickSwitchUtil.UnhookWinEvent(_hookDialog);   _hookDialog   = IntPtr.Zero; } } catch { }
            try { StopExplorerHook(); } catch { }

            QuickSwitchUtil.TrackedDialogs.Clear();
            _pendingPaths.Clear();
            while (QuickSwitchUtil.DialogEvents.TryDequeue(out _)) { }

            try { if (_staDispatcher != null) { _staDispatcher.InvokeShutdown(); _staDispatcher = null; } } catch { }
            try { if (_pumpTimer     != null) { _pumpTimer.Stop();               _pumpTimer     = null; } } catch { }
        }

        // ────────────────── Explorer Hook 惰性管理 (在 STA 线程调用) ──────────────────

        /// <summary>
        /// 开启 Explorer 导航 + 对话框焦点监控 hook.
        /// 在第一个文件对话框打开时由 OpenDialog() 调用.
        /// </summary>
        private static void StartExplorerHook()
        {
            if (_hookExplorer != IntPtr.Zero) return;

            // 范围: EVENT_SYSTEM_FOREGROUND (0x0003) ~ EVENT_OBJECT_NAMECHANGE (0x800C)
            //   - EVENT_SYSTEM_FOREGROUND : 窗口切换前台 → 用于"对话框重新获得焦点"检测
            //   - EVENT_OBJECT_NAMECHANGE : Explorer 标题栏变化 → 用于"窗口内导航"检测
            _hookExplorer = QuickSwitchUtil.SetWinEventHook(
                QuickSwitchUtil.EVENT_SYSTEM_FOREGROUND,
                QuickSwitchUtil.EVENT_OBJECT_NAMECHANGE,
                IntPtr.Zero,
                _explorerFocusProc,
                0, 0,
                QuickSwitchUtil.WINEVENT_OUTOFCONTEXT | QuickSwitchUtil.WINEVENT_SKIPOWNPROCESS);

            // 立即探测当前 Explorer 路径, 供对话框首次注入使用
            _staDispatcher?.BeginInvoke(new Action(() =>
            {
                try
                {
                    string p = QuickSwitchUtil.GetAnyExplorerPath();
                    if (!string.IsNullOrEmpty(p) && QuickSwitchUtil.IsValidPath(p))
                    {
                        QuickSwitchUtil.SetLastExplorerPath(p);
                        LogUtil.WriteQuickSwitchLog("StartExplorerHook: warmup path=" + p);
                    }
                    else
                    {
                        LogUtil.WriteQuickSwitchLog("StartExplorerHook: warmup got NULL");
                    }
                }
                catch (Exception ex)
                {
                    LogUtil.WriteQuickSwitchLog("StartExplorerHook warmup EX: " + ex.Message);
                }
            }), DispatcherPriority.Background);

            LogUtil.WriteQuickSwitchLog("StartExplorerHook: hook=" + _hookExplorer);
        }

        /// <summary>
        /// 停止 Explorer 监控 hook.
        /// 在最后一个文件对话框关闭时由 CloseDialog() 调用.
        /// </summary>
        private static void StopExplorerHook()
        {
            if (_hookExplorer == IntPtr.Zero) return;
            QuickSwitchUtil.UnhookWinEvent(_hookExplorer);
            _hookExplorer = IntPtr.Zero;
            LogUtil.WriteQuickSwitchLog("StopExplorerHook: 已停止");
        }

        // ─────────────────────────── WinEvent 回调 ───────────────────────────

        /// <summary>
        /// [始终开启] 对话框生命周期回调: 仅处理 #32770, 把事件压入队列.
        /// </summary>
        private static void DialogLifecycleCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hwnd == IntPtr.Zero) return;
            if (idObject != QuickSwitchUtil.OBJID_WINDOW) return;

            try
            {
                var sb = new System.Text.StringBuilder(64);
                QuickSwitchUtil.GetClassName(hwnd, sb, sb.Capacity);
                if (sb.ToString() != "#32770") return;

                LogUtil.WriteQuickSwitchLog("DialogLifecycle event=0x"
                    + eventType.ToString("X") + " hwnd=0x" + hwnd.ToString("X"));

                QuickSwitchUtil.DialogEvents.Enqueue(new QuickSwitchUtil.DialogEvent
                {
                    EventType = eventType,
                    Hwnd      = hwnd
                });
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("DialogLifecycleCallback EX: " + ex.Message);
            }
        }

        /// <summary>
        /// [惰性开启] Explorer 导航 + 对话框焦点回调.
        ///
        /// 情形 A — Explorer 窗口 (CabinetWClass / ExplorerWClass):
        ///   读取最新路径, 以"待注入"方式记录到各跟踪对话框, 不立即注入.
        ///
        /// 情形 B — 已跟踪的 #32770 获得前台焦点 (EVENT_SYSTEM_FOREGROUND):
        ///   将待注入路径应用到该对话框, 然后清空待注入.
        /// </summary>
        private static void ExplorerAndFocusCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hwnd == IntPtr.Zero) return;
            if (idObject != QuickSwitchUtil.OBJID_WINDOW) return;

            try
            {
                var sb = new System.Text.StringBuilder(64);
                QuickSwitchUtil.GetClassName(hwnd, sb, sb.Capacity);
                string cls = sb.ToString();

                // ── 情形 A: Explorer 窗口发生导航或切换前台 ──
                if (QuickSwitchUtil.IsExplorerClass(cls))
                {
                    if (_staDispatcher == null || QuickSwitchUtil.TrackedDialogs.IsEmpty) return;

                    _staDispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            string path = QuickSwitchUtil.GetPathFromExplorerWindow(hwnd);
                            if (string.IsNullOrEmpty(path) || !QuickSwitchUtil.IsValidPath(path)) return;

                            QuickSwitchUtil.SetLastExplorerPath(path);

                            // 只记录待注入路径, 不立即注入
                            foreach (var dlgHwnd in QuickSwitchUtil.TrackedDialogs.Keys)
                            {
                                if (!QuickSwitchUtil.IsWindow(dlgHwnd)) continue;
                                _pendingPaths[dlgHwnd] = path;
                                LogUtil.WriteQuickSwitchLog("ExplorerNav: pending recorded"
                                    + " dlg=0x" + dlgHwnd.ToString("X")
                                    + " event=0x" + eventType.ToString("X")
                                    + " path=" + path);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtil.WriteQuickSwitchLog("ExplorerAndFocusCallback[A] EX: " + ex.Message);
                        }
                    }), DispatcherPriority.Background);
                    return;
                }

                // ── 情形 B: 已跟踪的对话框重新获得前台焦点 ──
                if (cls == "#32770"
                    && eventType == QuickSwitchUtil.EVENT_SYSTEM_FOREGROUND
                    && QuickSwitchUtil.TrackedDialogs.ContainsKey(hwnd))
                {
                    // 原子地取走待注入路径; 若已被其他路径更新则放弃本次 (下次焦点再应用)
                    string pending;
                    if (!_pendingPaths.TryRemove(hwnd, out pending) || string.IsNullOrEmpty(pending))
                        return;

                    LogUtil.WriteQuickSwitchLog("DialogFocus: applying pending"
                        + " dlg=0x" + hwnd.ToString("X") + " path=" + pending);

                    // TryInject 内部使用 UI Automation, 必须在 STA 线程执行
                    var capturedHwnd = hwnd;
                    var capturedPath = pending;
                    _staDispatcher?.BeginInvoke(new Action(() =>
                    {
                        if (!QuickSwitchUtil.TrackedDialogs.ContainsKey(capturedHwnd)) return;
                        TryInject(capturedHwnd, capturedPath);
                        // 记录已注入的路径 (对话框可能在注入期间关闭, 用 ContainsKey 二次确认)
                        if (QuickSwitchUtil.TrackedDialogs.ContainsKey(capturedHwnd))
                            QuickSwitchUtil.TrackedDialogs[capturedHwnd] = capturedPath;
                    }), DispatcherPriority.Normal);   // Normal: 响应及时
                }
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("ExplorerAndFocusCallback outer EX: " + ex.Message);
            }
        }

        // ─────────────────────────── 对话框事件处理 (STA 线程) ───────────────────────────

        /// <summary>
        /// 在 STA dispatcher 上周期性消费对话框事件队列.
        /// </summary>
        public static void ProcessDialogEvents()
        {
            while (QuickSwitchUtil.DialogEvents.TryDequeue(out var evt))
            {
                HandleDialogEvent(evt);
            }
            CleanupClosedDialogs();
        }

        private static void HandleDialogEvent(QuickSwitchUtil.DialogEvent evt)
        {
            if (QuickSwitchUtil.ShouldSkipDialog(evt.Hwnd)) return;

            if (evt.EventType == QuickSwitchUtil.EVENT_OBJECT_CREATE
                || evt.EventType == QuickSwitchUtil.EVENT_OBJECT_SHOW)
            {
                OpenDialog(evt.Hwnd);
            }
            else if (evt.EventType == QuickSwitchUtil.EVENT_OBJECT_DESTROY
                || evt.EventType == QuickSwitchUtil.EVENT_OBJECT_HIDE)
            {
                CloseDialog(evt.Hwnd);
            }
        }

        /// <summary>
        /// 注册新打开的对话框, 惰性启动 Explorer hook, 并立即注入当前 Explorer 路径 (首次跟随).
        /// 在 STA dispatcher 上调用.
        /// </summary>
        private static void OpenDialog(IntPtr hwnd)
        {
            if (QuickSwitchUtil.TrackedDialogs.ContainsKey(hwnd)) return;

            bool wasEmpty = QuickSwitchUtil.TrackedDialogs.IsEmpty;
            QuickSwitchUtil.TrackedDialogs[hwnd] = "";

            // 第一个对话框打开 → 启动 Explorer 监控
            if (wasEmpty) StartExplorerHook();

            // 立即注入当前 Explorer 路径 (首次跟随, 用户体验优先)
            string initialPath = QuickSwitchUtil.GetLastExplorerPath();
            if (!string.IsNullOrEmpty(initialPath) && QuickSwitchUtil.IsValidPath(initialPath))
            {
                LogUtil.WriteQuickSwitchLog("OpenDialog: initial inject"
                    + " hwnd=0x" + hwnd.ToString("X") + " path=" + initialPath);
                TryInject(hwnd, initialPath);
                QuickSwitchUtil.TrackedDialogs[hwnd] = initialPath;
            }
            else
            {
                LogUtil.WriteQuickSwitchLog("OpenDialog: no initial path, hwnd=0x" + hwnd.ToString("X"));
            }
        }

        /// <summary>
        /// 移除已关闭的对话框, 并在所有对话框都关闭时停止 Explorer hook.
        /// 在 STA dispatcher 上调用.
        /// </summary>
        private static void CloseDialog(IntPtr hwnd)
        {
            if (!QuickSwitchUtil.TrackedDialogs.ContainsKey(hwnd)) return;

            QuickSwitchUtil.TrackedDialogs.TryRemove(hwnd, out _);
            _pendingPaths.TryRemove(hwnd, out _);

            // 所有对话框关闭 → 停止 Explorer 监控
            if (QuickSwitchUtil.TrackedDialogs.IsEmpty) StopExplorerHook();

            LogUtil.WriteQuickSwitchLog("CloseDialog: hwnd=0x" + hwnd.ToString("X")
                + " remaining=" + QuickSwitchUtil.TrackedDialogs.Count);
        }

        private static void CleanupClosedDialogs()
        {
            var hwnds = new System.Collections.Generic.List<IntPtr>(QuickSwitchUtil.TrackedDialogs.Keys);
            foreach (var hwnd in hwnds)
            {
                if (!QuickSwitchUtil.IsWindow(hwnd)) CloseDialog(hwnd);
            }
        }

        private static void TryInject(IntPtr hwnd, string path)
        {
            try { QuickSwitchUtil.TryInject(hwnd, path); }
            catch (Exception ex) { LogUtil.WriteErrorLog(ex, "QuickSwitch 注入异常"); }
        }
    }
}
