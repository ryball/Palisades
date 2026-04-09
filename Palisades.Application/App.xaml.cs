using Palisades.Helpers;
using Palisades.ViewModel;
using Sentry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace Palisades
{
    public partial class App : System.Windows.Application
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int SetCurrentProcessExplicitAppUserModelID(string appID);

        private const string AppUserModelId = "io.stouder.Palisades";
        private readonly HashSet<string> lastHiddenPalisadeIdentifiers = new(StringComparer.OrdinalIgnoreCase);
        private Forms.NotifyIcon? trayIcon;
        private Forms.ContextMenuStrip? trayMenu;

        public App()
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            SetCurrentProcessExplicitAppUserModelID(AppUserModelId);
            SetupSentry();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            SetupTrayIcon();

            PalisadesManager.LoadPalisades();
            if (PalisadesManager.palisades.Count == 0)
            {
                PalisadesManager.CreatePalisade();
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                ShowAllPalisades();
                ShowStartupBalloonTip();
            }), DispatcherPriority.ApplicationIdle);
        }

        private void SetupTrayIcon()
        {
            Drawing.Icon icon = Drawing.SystemIcons.Application;
            string iconPath = Path.Combine(AppContext.BaseDirectory, "Ressources", "icon.ico");
            if (File.Exists(iconPath))
            {
                icon = new Drawing.Icon(iconPath);
            }

            trayMenu = new Forms.ContextMenuStrip();
            trayMenu.Opening += (_, _) => RefreshTrayMenu();

            trayIcon = new Forms.NotifyIcon
            {
                Text = "Palisades",
                Icon = icon,
                ContextMenuStrip = trayMenu
            };

            RefreshTrayMenu();
            trayIcon.Visible = false;
            trayIcon.Visible = true;
            trayIcon.DoubleClick += (_, _) => Dispatcher.Invoke(ShowAllPalisades);
        }

        private void RefreshTrayMenu()
        {
            if (trayMenu == null)
            {
                return;
            }

            trayMenu.Items.Clear();
            trayMenu.Items.Add("Show all fences", null, (_, _) => Dispatcher.Invoke(ShowAllPalisades));
            trayMenu.Items.Add("Hide all fences", null, (_, _) => Dispatcher.Invoke(HideAllPalisades));
            trayMenu.Items.Add("Restore last hidden fences", null, (_, _) => Dispatcher.Invoke(RestoreLastHiddenPalisades));

            Forms.ToolStripMenuItem fencesMenu = new("Fences");
            foreach (KeyValuePair<string, View.Palisade> entry in PalisadesManager.palisades.OrderBy(entry => (entry.Value.DataContext as PalisadeViewModel)?.Name ?? "No name"))
            {
                string identifier = entry.Key;
                string fenceName = (entry.Value.DataContext as PalisadeViewModel)?.Name ?? "No name";
                bool isVisible = entry.Value.IsVisible;

                Forms.ToolStripMenuItem item = new(isVisible ? $"Hide {fenceName}" : $"Show {fenceName}");
                item.Click += (_, _) => Dispatcher.Invoke(() => TogglePalisadeVisibility(identifier));
                fencesMenu.DropDownItems.Add(item);
            }

            if (fencesMenu.DropDownItems.Count == 0)
            {
                fencesMenu.DropDownItems.Add("No fences").Enabled = false;
            }

            trayMenu.Items.Add(fencesMenu);
            trayMenu.Items.Add("New fence", null, (_, _) => Dispatcher.Invoke(PalisadesManager.CreatePalisade));
            trayMenu.Items.Add(new Forms.ToolStripSeparator());
            trayMenu.Items.Add("Exit", null, (_, _) => Dispatcher.Invoke(ExitApplication));
        }

        private void TogglePalisadeVisibility(string identifier)
        {
            if (!PalisadesManager.palisades.TryGetValue(identifier, out View.Palisade? palisade))
            {
                return;
            }

            if (palisade.IsVisible)
            {
                lastHiddenPalisadeIdentifiers.Add(identifier);
                palisade.Hide();
            }
            else
            {
                EnsurePalisadeIsVisible(palisade);

                if (!palisade.IsVisible)
                {
                    palisade.Show();
                }

                if (palisade.WindowState == WindowState.Minimized)
                {
                    palisade.WindowState = WindowState.Normal;
                }

                palisade.Activate();
                lastHiddenPalisadeIdentifiers.Remove(identifier);
            }

            RefreshTrayMenu();
        }

        private void ShowAllPalisades()
        {
            foreach (View.Palisade palisade in PalisadesManager.palisades.Values.ToList())
            {
                EnsurePalisadeIsVisible(palisade);

                if (!palisade.IsVisible)
                {
                    palisade.Show();
                }

                if (palisade.WindowState == WindowState.Minimized)
                {
                    palisade.WindowState = WindowState.Normal;
                }

                palisade.Activate();
            }

            lastHiddenPalisadeIdentifiers.Clear();
            RefreshTrayMenu();
        }

        private void HideAllPalisades()
        {
            lastHiddenPalisadeIdentifiers.Clear();

            foreach (KeyValuePair<string, View.Palisade> entry in PalisadesManager.palisades.ToList())
            {
                View.Palisade palisade = entry.Value;
                if (palisade.IsVisible)
                {
                    lastHiddenPalisadeIdentifiers.Add(entry.Key);
                }

                palisade.Hide();
            }

            RefreshTrayMenu();
        }

        private void RestoreLastHiddenPalisades()
        {
            if (lastHiddenPalisadeIdentifiers.Count == 0)
            {
                ShowAllPalisades();
                return;
            }

            foreach (string identifier in lastHiddenPalisadeIdentifiers.ToList())
            {
                if (!PalisadesManager.palisades.TryGetValue(identifier, out View.Palisade? palisade))
                {
                    continue;
                }

                EnsurePalisadeIsVisible(palisade);

                if (!palisade.IsVisible)
                {
                    palisade.Show();
                }

                if (palisade.WindowState == WindowState.Minimized)
                {
                    palisade.WindowState = WindowState.Normal;
                }

                palisade.Activate();
            }

            lastHiddenPalisadeIdentifiers.Clear();
            RefreshTrayMenu();
        }

        private static void EnsurePalisadeIsVisible(View.Palisade palisade)
        {
            Rect desktopBounds = new(SystemParameters.VirtualScreenLeft, SystemParameters.VirtualScreenTop, SystemParameters.VirtualScreenWidth, SystemParameters.VirtualScreenHeight);
            double width = palisade.Width > 0 ? palisade.Width : Math.Max(palisade.ActualWidth, 250);
            double height = palisade.Height > 0 ? palisade.Height : Math.Max(palisade.ActualHeight, 150);

            bool isOutsideVisibleArea =
                palisade.Left + 80 < desktopBounds.Left ||
                palisade.Top + 40 < desktopBounds.Top ||
                palisade.Left > desktopBounds.Right - 80 ||
                palisade.Top > desktopBounds.Bottom - 40;

            double minLeft = desktopBounds.Left + 20;
            double minTop = desktopBounds.Top + 20;
            double maxLeft = Math.Max(minLeft, desktopBounds.Right - width - 20);
            double maxTop = Math.Max(minTop, desktopBounds.Bottom - height - 20);

            if (isOutsideVisibleArea)
            {
                palisade.Left = minLeft;
                palisade.Top = minTop;
                return;
            }

            palisade.Left = Math.Min(Math.Max(palisade.Left, minLeft), maxLeft);
            palisade.Top = Math.Min(Math.Max(palisade.Top, minTop), maxTop);
        }

        private void ShowStartupBalloonTip()
        {
            if (trayIcon == null)
            {
                return;
            }

            trayIcon.BalloonTipTitle = "Palisades is running";
            trayIcon.BalloonTipText = "Use the tray icon to show, hide, restore, create, or exit your fences.";
            trayIcon.ShowBalloonTip(3000);
        }

        private void ExitApplication()
        {
            trayIcon?.Dispose();
            trayIcon = null;
            Shutdown();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            trayIcon?.Dispose();
            trayIcon = null;
            base.OnExit(e);
        }

        private void SetupSentry()
        {
            DispatcherUnhandledException += App_DispatcherUnhandledException;

            SentrySdk.Init(o =>
            {
                o.Dsn = "https://ffd9f3db270c4bd583ab3041d6264c38@o1336793.ingest.sentry.io/6605931";
                o.Debug = PEnv.IsDev();
                o.TracesSampleRate = 1;
            });
        }

        void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            SentrySdk.CaptureException(e.Exception);
            e.Handled = true;
        }
    }
}
