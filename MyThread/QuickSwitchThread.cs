using GeekDesk.Util;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Threading;

namespace GeekDesk.MyThread
{
    /// <summary>
    /// Quick Switch (文件夹跟随) 后台线程编排。
    ///
    /// 1. 启动 SetWinEventHook 监听全局窗口事件
    /// 2. 在专用 STA Dispatcher 上读取 UI Automation (Explorer / Dialog 地址栏)
    /// 3. 为每个被跟踪的 #32770 对话框启动一个轻量轮询线程, 实现"持续同步"
    /// </summary>
    public static class QuickSwitchThread
    {
        private static IntPtr _hookDialog;
        private static IntPtr _hookExplorer;
        private static Dispatcher _staDispatcher;

        private static readonly QuickSwitchUtil.WinEventDelegate _dialogProc = DialogEventCallback;
        private static readonly QuickSwitchUtil.WinEventDelegate _explorerProc = ExplorerEventCallback;

        // 跟踪轮询线程: HWND -> CancellationTokenSource
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<IntPtr, CancellationTokenSource> _pollers
            = new System.Collections.Concurrent.ConcurrentDictionary<IntPtr, CancellationTokenSource>();

        // 等待 Explorer 路径就绪后重试的对话框 (hwnd -> 重试剩余次数)
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<IntPtr, int> _pendingRetry
            = new System.Collections.Concurrent.ConcurrentDictionary<IntPtr, int>();

        // 防止在异步探测结果返回前, 反复向 STA 队列压入探测任务 (节流)
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<IntPtr, long> _lastProbeTick
            = new System.Collections.Concurrent.ConcurrentDictionary<IntPtr, long>();
        private const long ProbeThrottleTicks = 200 * TimeSpan.TicksPerMillisecond; // 200ms

        // Dispose 互斥
        private static readonly object _lock = new object();
        private static bool _running = false;

        // 周期性驱动 QuickSwitchUtil.DialogEvents 消费 (UI Automation 在 STA dispatcher 上执行)
        private static System.Windows.Threading.DispatcherTimer _pumpTimer;

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

                // Hook #32770 对话框的创建/销毁
                _hookDialog = QuickSwitchUtil.SetWinEventHook(
                    QuickSwitchUtil.EVENT_OBJECT_CREATE,
                    QuickSwitchUtil.EVENT_OBJECT_DESTROY,
                    IntPtr.Zero,
                    _dialogProc,
                    0, 0,
                    QuickSwitchUtil.WINEVENT_OUTOFCONTEXT | QuickSwitchUtil.WINEVENT_SKIPOWNPROCESS);

                // Hook 资源管理器位置变化 / 切换前台
                _hookExplorer = QuickSwitchUtil.SetWinEventHook(
                    QuickSwitchUtil.EVENT_SYSTEM_FOREGROUND,
                    QuickSwitchUtil.EVENT_OBJECT_LOCATIONCHANGE,
                    IntPtr.Zero,
                    _explorerProc,
                    0, 0,
                    QuickSwitchUtil.WINEVENT_OUTOFCONTEXT | QuickSwitchUtil.WINEVENT_SKIPOWNPROCESS);

                // 启动 pump: 在 STA dispatcher 上周期性处理对话框事件队列 (避免与 WPF UI 线程死锁)
                _pumpTimer = new System.Windows.Threading.DispatcherTimer(
                    System.Windows.Threading.DispatcherPriority.Background,
                    _staDispatcher);
                _pumpTimer.Interval = TimeSpan.FromMilliseconds(200);
                _pumpTimer.Tick += (s, e) => ProcessDialogEvents();
                _pumpTimer.Start();

                // 主动从现有 Explorer 窗口预热一次缓存, 避免首次对话框弹出时还要探测.
                if (_staDispatcher != null)
                {
                    _staDispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            string p = QuickSwitchUtil.GetAnyExplorerPath();
                            if (!string.IsNullOrEmpty(p) && QuickSwitchUtil.IsValidPath(p))
                            {
                                QuickSwitchUtil.SetLastExplorerPath(p);
                                LogUtil.WriteQuickSwitchLog("Warmup probe got path: " + p);
                            }
                            else
                            {
                                LogUtil.WriteQuickSwitchLog("Warmup probe got NULL path");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtil.WriteQuickSwitchLog("Warmup probe EX: " + ex.Message);
                        }
                    }), DispatcherPriority.Background);
                }

                LogUtil.WriteQuickSwitchLog("Hook 启动: selfPid=" + QuickSwitchUtil.GeekDeskPid
                    + " hookDialog=" + _hookDialog
                    + " hookExplorer=" + _hookExplorer
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

            try
            {
                if (_hookDialog != IntPtr.Zero)
                {
                    QuickSwitchUtil.UnhookWinEvent(_hookDialog);
                    _hookDialog = IntPtr.Zero;
                }
            }
            catch { }

            try
            {
                if (_hookExplorer != IntPtr.Zero)
                {
                    QuickSwitchUtil.UnhookWinEvent(_hookExplorer);
                    _hookExplorer = IntPtr.Zero;
                }
            }
            catch { }

            // 取消所有轮询线程
            foreach (var kv in _pollers)
            {
                try { kv.Value.Cancel(); } catch { }
            }
            _pollers.Clear();

            // 清空跟踪列表
            QuickSwitchUtil.TrackedDialogs.Clear();
            _pendingRetry.Clear();
            _lastProbeTick.Clear();
            while (QuickSwitchUtil.DialogEvents.TryDequeue(out _)) { }

            try
            {
                if (_staDispatcher != null)
                {
                    _staDispatcher.InvokeShutdown();
                    _staDispatcher = null;
                }
            }
            catch { }

            try
            {
                if (_pumpTimer != null)
                {
                    _pumpTimer.Stop();
                    _pumpTimer = null;
                }
            }
            catch { }
        }

        #region WinEvent 回调 (hook 线程)

        private static void DialogEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hwnd == IntPtr.Zero) return;
            if (idObject != QuickSwitchUtil.OBJID_WINDOW) return;

            try
            {
                string cls = null;
                var sb = new System.Text.StringBuilder(64);
                QuickSwitchUtil.GetClassName(hwnd, sb, sb.Capacity);
                cls = sb.ToString();

                if (cls != "#32770") return;

                LogUtil.WriteQuickSwitchLog("DialogEventCallback event=0x" + eventType.ToString("X") + " hwnd=0x" + hwnd.ToString("X"));

                QuickSwitchUtil.DialogEvents.Enqueue(new QuickSwitchUtil.DialogEvent
                {
                    EventType = eventType,
                    Hwnd = hwnd
                });
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("DialogEventCallback EX: " + ex.Message);
            }
        }

        private static void ExplorerEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hwnd == IntPtr.Zero) return;
            if (idObject != QuickSwitchUtil.OBJID_WINDOW) return;

            try
            {
                var sb = new System.Text.StringBuilder(64);
                QuickSwitchUtil.GetClassName(hwnd, sb, sb.Capacity);
                string cls = sb.ToString();

                if (!IsExplorerClass(cls)) return;

                // 在 STA 线程上读取 UI Automation
                if (_staDispatcher == null) return;

                _staDispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        string path = QuickSwitchUtil.GetPathFromExplorerWindow(hwnd);
                        if (!string.IsNullOrEmpty(path) && QuickSwitchUtil.IsValidPath(path))
                        {
                            QuickSwitchUtil.SetLastExplorerPath(path);
                            LogUtil.WriteQuickSwitchLog("ExplorerEvent event=0x" + eventType.ToString("X") + " cls=" + cls + " hwnd=0x" + hwnd.ToString("X") + " path=" + path);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtil.WriteQuickSwitchLog("ExplorerEvent EX: " + ex.Message);
                    }
                }), DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                LogUtil.WriteQuickSwitchLog("ExplorerEventCallback outer EX: " + ex.Message);
            }
        }

        private static bool IsExplorerClass(string cls)
        {
            return cls == "CabinetWClass"
                || cls == "ExplorerWClass"
                || cls == "Progman"
                || cls == "WorkerW";
        }

        #endregion

        /// <summary>
        /// 处理队列中累积的对话框事件。在 MainWindow UI 线程上周期性调用 (建议 100~200ms)。
        /// 由 QuickSwitchPump (UI 端) 调用, 保证 UI Automation 在 STA 上执行。
        /// </summary>
        public static void ProcessDialogEvents()
        {
            while (QuickSwitchUtil.DialogEvents.TryDequeue(out var evt))
            {
                HandleDialogEvent(evt);
            }

            // 处理等待重试的对话框 (Explorer 路径晚于对话框到达的场景)
            if (!_pendingRetry.IsEmpty)
            {
                var toRetry = new System.Collections.Generic.List<IntPtr>(_pendingRetry.Keys);
                foreach (var hwnd in toRetry)
                {
                    int left;
                    if (!_pendingRetry.TryGetValue(hwnd, out left)) continue;
                    if (!QuickSwitchUtil.IsWindow(hwnd))
                    {
                        _pendingRetry.TryRemove(hwnd, out _);
                        _lastProbeTick.TryRemove(hwnd, out _);
                        continue;
                    }
                    if (left <= 0)
                    {
                        _pendingRetry.TryRemove(hwnd, out _);
                        _lastProbeTick.TryRemove(hwnd, out _);
                        continue;
                    }
                    _pendingRetry[hwnd] = left - 1;
                    TryStartTracking(hwnd);
                }
            }

            // 清理已关闭的对话框
            CleanupClosedDialogs();
        }

        private static void HandleDialogEvent(QuickSwitchUtil.DialogEvent evt)
        {
            if (QuickSwitchUtil.ShouldSkipDialog(evt.Hwnd)) return;

            if (evt.EventType == QuickSwitchUtil.EVENT_OBJECT_CREATE
                || evt.EventType == QuickSwitchUtil.EVENT_OBJECT_SHOW)
            {
                TryStartTracking(evt.Hwnd);
            }
            else if (evt.EventType == QuickSwitchUtil.EVENT_OBJECT_DESTROY
                || evt.EventType == QuickSwitchUtil.EVENT_OBJECT_HIDE)
            {
                StopTracking(evt.Hwnd);
            }
        }

        private static void TryStartTracking(IntPtr hwnd)
        {
            if (QuickSwitchUtil.TrackedDialogs.ContainsKey(hwnd)) return;

            string targetPath = QuickSwitchUtil.GetLastExplorerPath();
            if (string.IsNullOrEmpty(targetPath))
            {
                // 缓存为空: 节流地向 STA dispatcher 异步探测任意 Explorer 路径, 不阻塞当前 pump
                if (_staDispatcher != null)
                {
                    long now = DateTime.UtcNow.Ticks;
                    long last;
                    if (!_lastProbeTick.TryGetValue(hwnd, out last) || (now - last) >= ProbeThrottleTicks)
                    {
                        _lastProbeTick[hwnd] = now;
                        LogUtil.WriteQuickSwitchLog("TryStartTracking: cache empty, probing Explorer path asynchronously for hwnd=0x" + hwnd.ToString("X"));
                        _staDispatcher.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                string p = QuickSwitchUtil.GetAnyExplorerPath();
                                if (!string.IsNullOrEmpty(p) && QuickSwitchUtil.IsValidPath(p))
                                {
                                    QuickSwitchUtil.SetLastExplorerPath(p);
                                    LogUtil.WriteQuickSwitchLog("Async probe got path: " + p);
                                    // 探测成功后立刻再尝试启动跟踪
                                    _staDispatcher.BeginInvoke(new Action(() => TryStartTracking(hwnd)),
                                        DispatcherPriority.Background);
                                }
                                else
                                {
                                    LogUtil.WriteQuickSwitchLog("Async probe got NULL path");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogUtil.WriteQuickSwitchLog("Async probe EX: " + ex.Message);
                            }
                        }), DispatcherPriority.Background);
                    }
                }
                // 加入重试队列, 下一轮 pump 会再次尝试 (探测成功或 _pendingRetry 耗尽时会被清掉)
                _pendingRetry[hwnd] = 50; // 50 * 200ms = 10s 重试窗口
                return;
            }

            if (!QuickSwitchUtil.IsValidPath(targetPath))
            {
                LogUtil.WriteQuickSwitchLog("TryStartTracking: invalid cached path: " + targetPath);
                return;
            }

            QuickSwitchUtil.TrackedDialogs[hwnd] = targetPath;
            _pendingRetry.TryRemove(hwnd, out _);
            LogUtil.WriteQuickSwitchLog("TryStartTracking: tracking hwnd=0x" + hwnd.ToString("X") + " path=" + targetPath);

            // 首次注入
            TryInject(hwnd, targetPath);

            // 启动轮询线程
            StartPoller(hwnd);
        }

        private static void StartPoller(IntPtr hwnd)
        {
            var cts = new CancellationTokenSource();
            _pollers[hwnd] = cts;

            var token = cts.Token;
            new System.Threading.Thread(() =>
            {
                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        System.Threading.Thread.Sleep(400);

                        if (token.IsCancellationRequested) break;
                        if (!QuickSwitchUtil.IsWindow(hwnd)) break;

                        string current = QuickSwitchUtil.GetLastExplorerPath();
                        if (string.IsNullOrEmpty(current)) continue;

                        string cached;
                        if (QuickSwitchUtil.TrackedDialogs.TryGetValue(hwnd, out cached)
                            && string.Equals(cached, current, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        TryInject(hwnd, current);
                        QuickSwitchUtil.TrackedDialogs[hwnd] = current;
                    }
                    catch
                    {
                        // ignored
                    }
                }

                _pollers.TryRemove(hwnd, out _);
                QuickSwitchUtil.TrackedDialogs.TryRemove(hwnd, out _);
            })
            { IsBackground = true }.Start();
        }

        private static void StopTracking(IntPtr hwnd)
        {
            CancellationTokenSource cts;
            if (_pollers.TryRemove(hwnd, out cts))
            {
                try { cts.Cancel(); cts.Dispose(); } catch { }
            }
            QuickSwitchUtil.TrackedDialogs.TryRemove(hwnd, out _);
            _pendingRetry.TryRemove(hwnd, out _);
            _lastProbeTick.TryRemove(hwnd, out _);
        }

        private static void CleanupClosedDialogs()
        {
            // 复制键列表避免在迭代中修改
            var hwnds = new System.Collections.Generic.List<IntPtr>(QuickSwitchUtil.TrackedDialogs.Keys);
            foreach (var hwnd in hwnds)
            {
                if (!QuickSwitchUtil.IsWindow(hwnd))
                {
                    StopTracking(hwnd);
                }
            }
        }

        private static void TryInject(IntPtr hwnd, string path)
        {
            try
            {
                QuickSwitchUtil.TryInject(hwnd, path);
            }
            catch (Exception ex)
            {
                LogUtil.WriteErrorLog(ex, "QuickSwitch 注入异常");
            }
        }
    }
}