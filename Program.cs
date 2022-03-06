using MuteWhenDeactivated;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UnmuteActive
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApplicationContext());
        }
    }

    public class TrayApplicationContext : ApplicationContext
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        static CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        static CancellationToken cancellationToken = cancellationTokenSource.Token;

        private NotifyIcon trayIcon;

        public TrayApplicationContext()
        {
            ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

            ToolStripItem exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += new EventHandler(Exit);

            contextMenuStrip.Items.AddRange(new ToolStripItem[] { exitItem });

            trayIcon = new NotifyIcon()
            {
                Text = "Unmute Active v0.1 (dev:camenduru)",
                Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath), // trayIcon is your NotifyIcon
            Visible = true
            };

            trayIcon.ContextMenuStrip = contextMenuStrip;

            Task task = Task.Run(async () =>
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    var processes = Process.GetProcesses().OrderBy(process => process.Id);
                    foreach (var process in processes)
                    {
                        var muteStatus = AudioManager.GetApplicationMute(process.Id);
                        if (muteStatus.HasValue)
                        {
                            IntPtr hwnd = GetForegroundWindow();
                            GetWindowThreadProcessId(hwnd, out uint pid);

                            var processForeground = Process.GetProcessById((int)pid);
                            if (process.ProcessName == processForeground.ProcessName)
                                AudioManager.SetApplicationMute(process.Id, false);
                            else
                                AudioManager.SetApplicationMute(process.Id, true);
                        }
                    }
                    await Task.Delay(1000);
                }
                cancellationToken.ThrowIfCancellationRequested();
            }, cancellationToken).ContinueWith(t =>
            {
                t.Exception?.Handle(e => true);
                Unmute();
            }, TaskContinuationOptions.OnlyOnCanceled);
        }

        void Exit(object sender, EventArgs e)
        {
            cancellationTokenSource.Cancel();
            Unmute();
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            trayIcon.Visible = false;
            Application.Exit();
        }

        void Unmute()
        {
            var processes = Process.GetProcesses().OrderBy(process => process.Id);
            foreach (var process in processes)
            {
                var muteStatus = AudioManager.GetApplicationMute(process.Id);
                if (muteStatus.HasValue)
                {
                    AudioManager.SetApplicationMute(process.Id, false);
                }
            }
        }
    }
}
