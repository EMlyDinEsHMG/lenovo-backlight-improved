using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LenovoBacklightImproved
{
    /*
     * Keyboard hook
     */
    using System.Runtime.InteropServices;
    internal static class KeyboardHook
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;


        private static IntPtr _hookID = IntPtr.Zero;
        private static LowLevelKeyboardProc _proc = HookCallback;

        public static event Action? OnKeyPressed;

        public static void Start()
        {
            _hookID = SetHook(_proc);
        }

        public static void Stop()
        {
            UnhookWindowsHookEx(_hookID);
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule!)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public KbdLlHookFlags flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [Flags]
        private enum KbdLlHookFlags : uint
        {
            LLKHF_EXTENDED = 0x01,
            LLKHF_INJECTED = 0x10,
            LLKHF_ALTDOWN = 0x20,
            LLKHF_UP = 0x80
        }

        private static HashSet<int> keysCurrentlyDown = new HashSet<int>();
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);

                // Ignore injected (synthetic) input
                if ((hookStruct.flags & KbdLlHookFlags.LLKHF_INJECTED) != 0)
                {
                    return CallNextHookEx(_hookID, nCode, wParam, lParam);
                }

                int vkCode = Marshal.ReadInt32(lParam);

                if (wParam == (IntPtr)WM_KEYDOWN)
                {
                    // Detect new key press (not repeat)
                    if (!keysCurrentlyDown.Contains(vkCode))
                    {
                        Trace.WriteLine($"Key down - {vkCode}");
                        keysCurrentlyDown.Add(vkCode);
                        OnKeyPressed?.Invoke();
                    }
                }
                else if (wParam == (IntPtr)WM_KEYUP)
                {
                    // Remove key from set
                    keysCurrentlyDown.Remove(vkCode);
                }
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        // P/Invoke
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }

    class Program : Form
    {
        /*
         *  Internal Constants
         */
        internal const string appNameSpaced = "Lenovo Backlight Improved";
        internal static readonly string appName = appNameSpaced.Replace(" ", "");

        /*
         *  Private Constants
         */
        private const string defaultDllLocation = "C:\\ProgramData\\Lenovo";
        private readonly string[] defaultDlls = ["C:\\ProgramData\\Lenovo\\Vantage\\Addins\\ThinkKeyboardAddin\\1.0.0.18\\Keyboard_Core.dll",
                                                 "C:\\ProgramData\\Lenovo\\ImController\\Plugins\\ThinkKeyboardPlugin\\x86\\Keyboard_Core.dll"];

        /*
         *  Private Variabled
         */
        private string lastDllTried = "";

        private SystemTray? systemTray;

        private KeyboardBacklightStatusControl? keyboardBacklightStatusControl;

        private System.Timers.Timer timeoutTimer = new System.Timers.Timer();
        private uint inactivityTimeout;

        private byte selectedBacklightLevel;
        private byte currentBacklightStatus;

        private bool isStatusChangeFromUs = false;

        private bool ignoreMouseEvents = false;

        DateTime lastBacklightOffTime = DateTime.MinValue;

        DateTime lastKeyPressTime = DateTime.Now;


        /*
         *  Power Setting Variables
         */
        private static Guid GUID_SYSTEM_POWER_SETTING = new Guid("45BC44C4-4E13-4F23-9C0E-9F7593BEFCA0");

        private readonly Message aboutToSleep = new Message
        {
            Msg = 0x320,
            WParam = unchecked((nint)0xe3005ba3),
        };

        private readonly Message resumeFromSleep = new Message
        {
            Msg = 0x218,
            WParam = 0x07,
        };

        private void OnApplicationExit(object? sender, EventArgs e)
        {
            Trace.WriteLine("Application exiting - stopping keyboard hook.");
            KeyboardHook.Stop();
        }

        /*
         *  Constructors
         */
        private Program()
        {
            try
            {
                loadDll(Properties.Settings.Default.DllPath);
            }
            catch (Exception ex)
            {
                string? defaultDll = defaultDlls.FirstOrDefault(dllPath => File.Exists(dllPath));

                if (defaultDll == null)
                {
                    defaultDll = FindFile(defaultDllLocation, "Keyboard_Core.dll");
                }

                if (defaultDll != null && Properties.Settings.Default.DllPath != defaultDll)
                {
                    try
                    {
                        loadDll(defaultDll);
                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show($"Configured DLL file could not be loaded: '{ex.Message}'\n" +
                                        $"Default DLL file also could not be loaded: '{ex2.Message}'", $"{appNameSpaced} - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show(ex.Message, $"{appNameSpaced} - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Register the power setting notification
            RegisterPowerSettingNotification(this.Handle, ref GUID_SYSTEM_POWER_SETTING, 0);

            systemTray = new SystemTray(this);

            inactivityTimeout = Properties.Settings.Default.InactivityTimeout;
            timeoutTimer.Interval = Properties.Settings.Default.CheckInterval;

            selectBacklightLevel(Properties.Settings.Default.BacklightLevel);

            // Setup and start keyboard activity monitor
            KeyboardHook.OnKeyPressed += () =>
            {
                Trace.WriteLine("Key press detected.");
                lastKeyPressTime = DateTime.Now;

                if (selectedBacklightLevel > 2 && currentBacklightStatus == 0)
                {
                    Trace.WriteLine("Key press detected - turning backlight on.");
                    setBacklightStatusSafe(keyboardBacklightStatusControl, getStatusFromLevel(selectedBacklightLevel));
                }
            };

            KeyboardHook.Start();

            // To Clean up
            Application.ApplicationExit += OnApplicationExit;

            Application.Run(systemTray);
        }

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool createdNew = true;
            using (Mutex mutex = new Mutex(true, appName, out createdNew))
            {
                if (!createdNew)
                {
                    MessageBox.Show("Application is already running! Check the system tray!", appNameSpaced, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ApplicationConfiguration.Initialize();
                Application.Run(new Program());
            }
        }

        /*
         *  Miscellaneous Functions
         */
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr RegisterPowerSettingNotification(IntPtr hRecipient, ref Guid PowerSettingGuid, int Flags);

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == aboutToSleep.Msg && m.WParam == aboutToSleep.WParam)
            {
                Trace.WriteLine("The system is going to sleep! Stopping timeoutTimer!");
                stopTimer();
                KeyboardHook.Stop();
            }
            else if (m.Msg == resumeFromSleep.Msg && m.WParam == resumeFromSleep.WParam)
            {
                Trace.WriteLine("The system has waken up! Starting timeoutTimer!");
                setBacklightStatusSafe(keyboardBacklightStatusControl, getStatusFromLevel(selectedBacklightLevel));
                startTimer();
                KeyboardHook.Start();
            }

            // Call the base method
            base.WndProc(ref m);
        }
        
        /*
         *  Internal Helper Functions
         */
        internal void selectBacklightLevel(byte newBacklightLevel)
        {
            selectedBacklightLevel = newBacklightLevel;
            Properties.Settings.Default.BacklightLevel = selectedBacklightLevel;
            Properties.Settings.Default.Save();
            
            if (systemTray != null)
            {
                systemTray.updateLevelsChecked();
            }
        }

        internal byte getSelectedBacklightLevel()
        {
            return selectedBacklightLevel;
        }

        /*
         *  Private Helper Functions
         */
        private byte getStatusFromLevel(byte level)
        {
            return level > 2 ? (byte)(level % 3 + 1) : level;
        }

        private void startTimer()
        {
            SystemIdle.makeNotIdle();
            timeoutTimer.Start();
        }

        private void stopTimer()
        {
            timeoutTimer.Stop();
        }

        private void setBacklightStatusSafe(KeyboardBacklightStatusControl? keyboardBacklightStatusControl, byte expectedStatus)
        {
            if (keyboardBacklightStatusControl != null)
            {
                keyboardBacklightStatusControl.SetStatus(expectedStatus);
                currentBacklightStatus = expectedStatus;
                isStatusChangeFromUs = true;
            }
        }

        private void checkAndCorrectBacklightStatus(object? sender, System.Timers.ElapsedEventArgs e, KeyboardBacklightStatusControl keyboardBacklightStatusControl)
        {
            // Track keyboard key press
            TimeSpan idleDuration = DateTime.Now - lastKeyPressTime;

            // Track keyboard and mouse events 
            uint idleTime = SystemIdle.getIdleTime();

            // Check status 
            byte newBacklightStatus = keyboardBacklightStatusControl.GetStatus();
            byte expectedStatus = getStatusFromLevel(selectedBacklightLevel);
            TimeSpan ignoreIdleResetDuration = TimeSpan.FromMilliseconds(getTimerInterval());

            Trace.WriteLine($"Idle time: {idleTime}");
            Trace.WriteLine($"Idle duration for key press: {idleDuration}");
            Trace.WriteLine($"Backlight status check: {currentBacklightStatus}, {newBacklightStatus}, {expectedStatus}, {isStatusChangeFromUs}, {getTimerInterval()}");

            if (!ignoreMouseEvents)
            {
                if (newBacklightStatus != currentBacklightStatus && isStatusChangeFromUs == false)
                {
                    selectBacklightLevel((byte)((selectedBacklightLevel + 1) % 5));
                    SystemIdle.makeNotIdle();

                    Trace.WriteLine($"Backlight level change: {selectedBacklightLevel}");

                    expectedStatus = getStatusFromLevel(selectedBacklightLevel);

                    if (expectedStatus != newBacklightStatus)
                    {
                        setBacklightStatusSafe(keyboardBacklightStatusControl, expectedStatus);
                        Trace.WriteLine($"Change needed!");
                    }
                    else
                    {
                        currentBacklightStatus = expectedStatus;
                        Trace.WriteLine($"No change needed!");
                    }
                }
                else if ((DateTime.Now - lastBacklightOffTime) < ignoreIdleResetDuration && isStatusChangeFromUs)
                {
                    // Ignore the idle time reset for this short duration
                    idleTime = inactivityTimeout + 1; // Force idle to "long enough"
                }
                else if (isStatusChangeFromUs && idleTime < getTimerInterval())
                {
                    isStatusChangeFromUs = false;
                }

                Trace.WriteLine($"Current backlight level: {selectedBacklightLevel}");

                if (selectedBacklightLevel > 2)
                {
                    if (idleTime >= inactivityTimeout && newBacklightStatus != 0)
                    {
                        setBacklightStatusSafe(keyboardBacklightStatusControl, 0);
                        lastBacklightOffTime = DateTime.Now;  // Track when backlight was turned off
                        Trace.WriteLine($"Idle! Shutting off!");
                    }
                    else if (isStatusChangeFromUs == false && idleTime < inactivityTimeout && newBacklightStatus != getStatusFromLevel(selectedBacklightLevel))
                    {
                        setBacklightStatusSafe(keyboardBacklightStatusControl, expectedStatus);
                        Trace.WriteLine($"Not idle! Setting to {expectedStatus}");
                    }
                }
                else if (expectedStatus != newBacklightStatus)
                {
                    Trace.WriteLine($"Setting to expected {expectedStatus}");
                    setBacklightStatusSafe(keyboardBacklightStatusControl, getStatusFromLevel(selectedBacklightLevel));
                }
            }
            else
            {
                if (selectedBacklightLevel > 2)
                {
                    if (idleDuration.TotalMilliseconds >= inactivityTimeout && newBacklightStatus != 0)
                    {
                        setBacklightStatusSafe(keyboardBacklightStatusControl, 0);
                        lastBacklightOffTime = DateTime.Now;
                        Trace.WriteLine($"Idle for {idleDuration.TotalMilliseconds}ms. Turning off backlight.");
                    }
                    else if (isStatusChangeFromUs == false && idleDuration.TotalMilliseconds < inactivityTimeout && newBacklightStatus != getStatusFromLevel(selectedBacklightLevel))
                    {
                        setBacklightStatusSafe(keyboardBacklightStatusControl, expectedStatus);
                        Trace.WriteLine("User active. Turning on backlight.");
                    }
                }
            }
        }
        private static string? FindFile(string path, string fileSearched)
        {
            try
            {
                string? foundFile = Directory.GetFiles(path).FirstOrDefault(file => file.EndsWith(fileSearched));

                if (foundFile != null)
                {
                    return foundFile;
                }

                string[] dirs = Directory.GetDirectories(path);

                foreach (string dir in dirs)
                {
                    foundFile = FindFile(dir, fileSearched);
                    
                    if (foundFile != null)
                    {
                        return foundFile;
                    }
                }

            }
            catch (UnauthorizedAccessException ex)
            {
                // No access to folder! Move on!
            }

            return null;
        }

        /*
         *  Public Getter Functions
         */
        public uint getInactivityTimeout()
        {
            return inactivityTimeout;
        }

        public uint getTimerInterval()
        {
            return (uint)timeoutTimer.Interval;
        }

        public string getDllPath()
        {
            if (keyboardBacklightStatusControl != null)
            {
                return keyboardBacklightStatusControl.getDll();
            }

            return "No DLL loaded!";
        }

        public string getLastDllTried()
        {
            return lastDllTried;
        }

        /*
         *  Public Setter Functions
         */

        public void setIgnoreMouseEvents(bool status)
        {
            ignoreMouseEvents = status;
        }

        public void setInactivityTimeout(uint newTimeout)
        {
            inactivityTimeout = newTimeout;
        }

        public void setTimerInterval(uint newInterval)
        {
            stopTimer();
            timeoutTimer.Interval = newInterval;
            startTimer();
        }

        public void loadDll(string dllPath)
        {
            if (timeoutTimer.Enabled == true)
            {
                stopTimer();
            }

            lastDllTried = dllPath;
            keyboardBacklightStatusControl = new KeyboardBacklightStatusControl(dllPath);

            setBacklightStatusSafe(keyboardBacklightStatusControl, getStatusFromLevel(selectedBacklightLevel));

            timeoutTimer.Elapsed += (sender, e) => checkAndCorrectBacklightStatus(sender, e, keyboardBacklightStatusControl);
            startTimer();
        }
    }
}