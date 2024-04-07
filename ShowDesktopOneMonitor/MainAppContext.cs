using FrigoTab;
using ShowDesktopOneMonitor.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShowDesktopOneMonitor
{
    public class MainAppContext : ApplicationContext
    {
        private NotifyIcon trayIcon = null;
        private List<DesktopWindowID>[] PrevStateByScreen = new List<DesktopWindowID>[0];


        public static void GetAutoRun(ref bool isAutoRun)
        {
            string path = System.Windows.Forms.Application.ExecutablePath;
            Microsoft.Win32.RegistryKey rk0 = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(@"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run");
            object value = rk0.GetValue("App");

            if (value != null && value.ToString().Length > 0)
                isAutoRun = true;
            else
                isAutoRun = false;
        }

        public static void SetAutoRun(bool isAutoRun)
        {
            string path = System.Windows.Forms.Application.ExecutablePath;
            //设置是否自动启动
            if (isAutoRun)
            {
                Microsoft.Win32.RegistryKey rk1 = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(@"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run");
                rk1.SetValue("App", @"""" + path + @"""");
                rk1.Close();
            }
            else
            {
                Microsoft.Win32.RegistryKey rk2 = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(@"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run");
                rk2.DeleteValue("App", false);
                rk2.Close();
            }
        }


        public MainAppContext()
        {
            Application.ThreadException += this.Application_ThreadException;
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            AppDomain.CurrentDomain.UnhandledException += this.CurrentDomain_UnhandledException;

            bool autoRun = false;
            GetAutoRun(ref autoRun);
            trayIcon = new NotifyIcon()
            {
                Icon = Resources.sde,
                ContextMenu = new ContextMenu(new MenuItem[] {
                    new MenuItem("auto start option", (s, e) => {
                    var menuItem = s as MenuItem;
                    menuItem.Checked = !menuItem.Checked; // 切换勾选状态
                    SetAutoRun( menuItem.Checked);
                }) { Checked = autoRun }, // 初始未勾选
                    new MenuItem("Exit", (s, e) => {trayIcon.Visible = false; Application.Exit(); })
                }),
                Visible = true,
                Text = "Show Desktop Enhanced",
            };
            PrevStateByScreen = new List<DesktopWindowID>[Screen.AllScreens.Length];

            Keys hotKey = SettingsManager.ReadHotkey();
            KeyModifiers keyModifiers = SettingsManager.ReadKeyModifiers();

            HotKeyManager.RegisterHotKey(hotKey, keyModifiers);
            HotKeyManager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(OnHotkeyPressed);


            SetAutoRun(true);
        }

        private void OnHotkeyPressed(object sender, HotKeyEventArgs e)
        {
            OnShowDesktopKeyComb();
        }

        private void OnShowDesktopKeyComb()
        {
            Console.WriteLine("Toogling hide windows on screen...");

            // 1. Get current screen
            Screen currentScreen = Screen.FromPoint(Cursor.Position);
            int screenIdx = Array.IndexOf(Screen.AllScreens, currentScreen);

            // 2. Get windows on selected screen
            List<WindowHandle> windows = GetWindowsOnScreen(currentScreen);

            // 3. Has the list changed from the previous list ?
            List<DesktopWindowID> newWindowIDs = ConvertWindowsToIDs(windows);

            // screen number may have changed
            Array.Resize(ref PrevStateByScreen, Screen.AllScreens.Length);

            // restore if all windows are minimized AND prev state differs only by windows style
            if (newWindowIDs.All(x => x.WindowStyle != WindowStyles.Visible) && PrevStateByScreen[screenIdx] != null
                                        && DoesPrevStateDiffersOnlyByWindowsStyle(newWindowIDs, screenIdx))
            {
                restoreAllWindows(screenIdx);
            }
            else
            {
                minimizeAllWindows(newWindowIDs, screenIdx);
            }
        }
        private void minimizeAllWindows(List<DesktopWindowID> windowList, int screenIdx)
        {
            //int count = 0;
            // sort by ZOrder to restore windows in reverse order
            windowList = windowList.Select(x => new { window = x, zOrder = WindowApi.GetWindowZOrder(x.WindowHandle) })
                                                            .OrderByDescending(x => x.zOrder).Select(x => x.window).ToList();
            foreach (var window in windowList)
            {
                if (window.WindowStyle == WindowStyles.Visible)
                {
                    window.SourceHandleObj.SetMinimizeWindow();

                    //count++;
                }
            }

            PrevStateByScreen[screenIdx] = windowList;
            //Console.WriteLine($"Minimizing {count} windows!");
        }
        private void restoreAllWindows(int screenIdx)
        {
            //int count = 0;
            if (PrevStateByScreen[screenIdx] != null)
            {
                foreach (var window in PrevStateByScreen[screenIdx].Reverse<DesktopWindowID>())
                {
                    if (window.WindowStyle == WindowStyles.Visible)
                    {
                        window.SourceHandleObj.SetRestoreWindow();

                        //count++;
                    }
                }
            }

            PrevStateByScreen[screenIdx] = null;
            //Console.WriteLine($"Restoring {count} windows!");
        }

        private List<WindowHandle> GetWindowsOnScreen(Screen screen)
        {
            WindowFinder finder = new WindowFinder();

            var windows = finder.Windows.Where(x => x.GetScreen().Equals(screen)).ToList();
            return windows;
        }
        private List<DesktopWindowID> ConvertWindowsToIDs(List<WindowHandle> windows)
        {
            List<DesktopWindowID> list = new List<DesktopWindowID>(windows.Count);
            for (int i = 0; i < windows.Count; i++)
            {
                var window = windows[i];
                list.Add(new DesktopWindowID(window));
            }
            return list;
        }
        private bool DoesPrevStateDiffersOnlyByWindowsStyle(List<DesktopWindowID> newList, int screenIdx)
        {
            if (PrevStateByScreen[screenIdx] == null) return false; // prev state is null
            if (newList.Count != PrevStateByScreen[screenIdx].Count) return false; // count differs
            if (false == newList.All(x => PrevStateByScreen[screenIdx].Contains(x))) return false; // windows differ

            // return true if style differs
            return false == newList.All(x => PrevStateByScreen[screenIdx].First(y => y == x).WindowStyle.Equals(x.WindowStyle));
        }

        ~MainAppContext() //Destructor
        {
            SettingsManager.Save();
        }

        private void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show("Необработанное исключение: " + e.Exception.ToString(), "Show Desktop Enhanced", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show($"Необработанное исключение: {e.ExceptionObject as Exception}", "Show Desktop Enhanced", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    public class DesktopWindowID : IEquatable<DesktopWindowID>
    {
        public WindowHandle SourceHandleObj;

        public IntPtr WindowHandle = IntPtr.Zero;
        public WindowStyles WindowStyle = WindowStyles.Disabled;

        public DesktopWindowID(WindowHandle sourceHandleObj)
        {
            this.SourceHandleObj = sourceHandleObj;

            this.WindowHandle = this.SourceHandleObj.GetHandle();

            this.WindowStyle = WindowStyles.Disabled;
            var wndStyle = this.SourceHandleObj.GetWindowStyles();

            if (wndStyle.HasFlag(WindowStyles.Minimize))
                this.WindowStyle = WindowStyles.Minimize;
            else if (wndStyle.HasFlag(WindowStyles.Visible))
                this.WindowStyle = WindowStyles.Visible;
        }

        public bool Equals(DesktopWindowID other)
        {
            return Equals((object)other);
        }
        public override int GetHashCode()
        {
            return this.WindowHandle.GetHashCode();
        }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            return this.GetHashCode() == obj.GetHashCode();
        }
        public static bool operator ==(DesktopWindowID obj1, DesktopWindowID obj2)
        {
            if (ReferenceEquals(obj1, null))
            {
                return ReferenceEquals(obj2, null);
            }
            return obj1.Equals(obj2);
        }
        public static bool operator !=(DesktopWindowID obj1, DesktopWindowID obj2)
        {
            return (obj1 == obj2) == false;
        }
    }
}