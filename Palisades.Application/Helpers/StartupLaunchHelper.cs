using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;

namespace Palisades.Helpers
{
    internal static class StartupLaunchHelper
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

        public static bool IsEnabled()
        {
            using RegistryKey? runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
            string? launchCommand = runKey?.GetValue(AppBranding.StartupValueName) as string;
            string? legacyLaunchCommand = runKey?.GetValue(AppBranding.LegacyName) as string;
            return !string.IsNullOrWhiteSpace(launchCommand)
                || !string.IsNullOrWhiteSpace(legacyLaunchCommand)
                || HasLegacyStartupShortcut();
        }

        public static void SetEnabled(bool enabled)
        {
            using RegistryKey runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, true) ?? Registry.CurrentUser.CreateSubKey(RunKeyPath);

            if (enabled)
            {
                string executablePath = GetExecutablePath();
                if (!string.IsNullOrWhiteSpace(executablePath) && File.Exists(executablePath))
                {
                    runKey.SetValue(AppBranding.StartupValueName, $"\"{executablePath}\"");
                    runKey.DeleteValue(AppBranding.LegacyName, false);
                }
            }
            else
            {
                runKey.DeleteValue(AppBranding.StartupValueName, false);
                runKey.DeleteValue(AppBranding.LegacyName, false);
                RemoveLegacyStartupShortcuts();
            }
        }

        private static string GetExecutablePath()
        {
            string? processPath = Environment.ProcessPath;
            if (!string.IsNullOrWhiteSpace(processPath))
            {
                return processPath;
            }

            using Process currentProcess = Process.GetCurrentProcess();
            return currentProcess.MainModule?.FileName ?? string.Empty;
        }

        private static bool HasLegacyStartupShortcut()
        {
            string startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            return File.Exists(Path.Combine(startupFolder, AppBranding.DisplayName + ".lnk"))
                || File.Exists(Path.Combine(startupFolder, AppBranding.LegacyName + ".lnk"));
        }

        private static void RemoveLegacyStartupShortcuts()
        {
            string startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            foreach (string shortcutName in new[] { AppBranding.DisplayName + ".lnk", AppBranding.LegacyName + ".lnk" })
            {
                string shortcutPath = Path.Combine(startupFolder, shortcutName);
                if (File.Exists(shortcutPath))
                {
                    try
                    {
                        File.Delete(shortcutPath);
                    }
                    catch
                    {
                        // Best-effort cleanup only.
                    }
                }
            }
        }
    }
}
