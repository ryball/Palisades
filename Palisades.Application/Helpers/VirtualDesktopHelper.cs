using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Palisades.Helpers
{
    public sealed class VirtualDesktopInfo
    {
        public string DesktopId { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public bool IsCurrent { get; init; }
        public bool IsSelected { get; init; }
        public bool CanMoveTo { get { return !IsCurrent; } }
        public string DisplayName { get { return IsCurrent ? $"{Name} (current)" : Name; } }
    }

    internal static class VirtualDesktopHelper
    {
        private const string VirtualDesktopsKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops";
        private const string VirtualDesktopsSubKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops\Desktops";
        private const string VirtualDesktopIdsValueName = "VirtualDesktopIDs";
        private const string CurrentVirtualDesktopValueName = "CurrentVirtualDesktop";

        public static bool HasMultipleDesktops()
        {
            return GetDesktopIds().Count > 1;
        }

        public static IReadOnlyList<VirtualDesktopInfo> GetDesktops()
        {
            List<Guid> desktopIds = GetDesktopIds();
            Dictionary<Guid, string> desktopNames = GetDesktopNames();
            Guid currentDesktopId = GetCurrentDesktopId();
            List<VirtualDesktopInfo> desktops = new();

            for (int index = 0; index < desktopIds.Count; index++)
            {
                Guid desktopId = desktopIds[index];
                string fallbackName = $"Desktop {index + 1}";
                string name = desktopNames.TryGetValue(desktopId, out string? storedName) && !string.IsNullOrWhiteSpace(storedName)
                    ? storedName.Trim()
                    : fallbackName;

                desktops.Add(new VirtualDesktopInfo
                {
                    DesktopId = desktopId.ToString(),
                    Name = name,
                    IsCurrent = desktopId == currentDesktopId
                });
            }

            return desktops;
        }

        public static bool TryMoveWindowToDesktop(IntPtr windowHandle, Guid targetDesktopId)
        {
            if (windowHandle == IntPtr.Zero || targetDesktopId == Guid.Empty)
            {
                return false;
            }

            try
            {
                IVirtualDesktopManager manager = (IVirtualDesktopManager)new CVirtualDesktopManager();
                return manager.MoveWindowToDesktop(windowHandle, targetDesktopId) >= 0;
            }
            catch
            {
                return false;
            }
        }

        public static bool TryMoveWindowToPreviousDesktop(IntPtr windowHandle)
        {
            return TryMoveWindowToAdjacentDesktop(windowHandle, -1);
        }

        public static bool TryMoveWindowToNextDesktop(IntPtr windowHandle)
        {
            return TryMoveWindowToAdjacentDesktop(windowHandle, 1);
        }

        private static bool TryMoveWindowToAdjacentDesktop(IntPtr windowHandle, int offset)
        {
            if (windowHandle == IntPtr.Zero)
            {
                return false;
            }

            List<Guid> desktopIds = GetDesktopIds();
            if (desktopIds.Count <= 1)
            {
                return false;
            }

            if (!TryGetWindowDesktopId(windowHandle, out Guid currentDesktopId))
            {
                return false;
            }

            int currentIndex = desktopIds.FindIndex(id => id == currentDesktopId);
            if (currentIndex < 0)
            {
                return false;
            }

            int targetIndex = currentIndex + Math.Sign(offset);
            if (targetIndex < 0 || targetIndex >= desktopIds.Count)
            {
                return false;
            }

            Guid targetDesktopId = desktopIds[targetIndex];
            return TryMoveWindowToDesktop(windowHandle, targetDesktopId);
        }

        private static bool TryGetWindowDesktopId(IntPtr windowHandle, out Guid desktopId)
        {
            desktopId = Guid.Empty;

            try
            {
                IVirtualDesktopManager manager = (IVirtualDesktopManager)new CVirtualDesktopManager();
                return manager.GetWindowDesktopId(windowHandle, out desktopId) >= 0 && desktopId != Guid.Empty;
            }
            catch
            {
                return false;
            }
        }

        public static Guid GetCurrentDesktopId()
        {
            try
            {
                using RegistryKey? key = Registry.CurrentUser.OpenSubKey(VirtualDesktopsKeyPath);
                if (key?.GetValue(CurrentVirtualDesktopValueName) is byte[] rawCurrentDesktop && rawCurrentDesktop.Length >= 16)
                {
                    byte[] guidBytes = new byte[16];
                    Buffer.BlockCopy(rawCurrentDesktop, 0, guidBytes, 0, 16);
                    return new Guid(guidBytes);
                }
            }
            catch
            {
                // Best effort only.
            }

            return Guid.Empty;
        }

        public static string GetCurrentDesktopIdString()
        {
            Guid currentDesktopId = GetCurrentDesktopId();
            return currentDesktopId == Guid.Empty ? string.Empty : currentDesktopId.ToString();
        }

        private static Dictionary<Guid, string> GetDesktopNames()
        {
            Dictionary<Guid, string> desktopNames = new();

            try
            {
                using RegistryKey? key = Registry.CurrentUser.OpenSubKey(VirtualDesktopsSubKeyPath);
                if (key == null)
                {
                    return desktopNames;
                }

                foreach (string subKeyName in key.GetSubKeyNames())
                {
                    if (!Guid.TryParse(subKeyName, out Guid desktopId))
                    {
                        continue;
                    }

                    using RegistryKey? desktopKey = key.OpenSubKey(subKeyName);
                    if (desktopKey == null)
                    {
                        continue;
                    }

                    string? name = desktopKey.GetValue("Name") as string;
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        name = desktopKey.GetValue("VirtualDesktopName") as string;
                    }

                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        desktopNames[desktopId] = name.Trim();
                    }
                }
            }
            catch
            {
                // Virtual desktop names are best effort only.
            }

            return desktopNames;
        }

        private static List<Guid> GetDesktopIds()
        {
            List<Guid> desktopIds = new();

            try
            {
                using RegistryKey? key = Registry.CurrentUser.OpenSubKey(VirtualDesktopsKeyPath);
                if (key?.GetValue(VirtualDesktopIdsValueName) is not byte[] rawIds || rawIds.Length < 16)
                {
                    return desktopIds;
                }

                for (int index = 0; index + 16 <= rawIds.Length; index += 16)
                {
                    byte[] guidBytes = new byte[16];
                    Buffer.BlockCopy(rawIds, index, guidBytes, 0, 16);
                    desktopIds.Add(new Guid(guidBytes));
                }
            }
            catch
            {
                // Virtual desktop enumeration is best effort only.
            }

            return desktopIds;
        }

        [ComImport]
        [Guid("a5cd92ff-29be-454c-8d04-d82879fb3f1b")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IVirtualDesktopManager
        {
            int IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow, out int onCurrentDesktop);
            int GetWindowDesktopId(IntPtr topLevelWindow, out Guid desktopId);
            int MoveWindowToDesktop(IntPtr topLevelWindow, [MarshalAs(UnmanagedType.LPStruct)] Guid desktopId);
        }

        [ComImport]
        [Guid("aa509086-5ca9-4c25-8f95-589d3c07b48a")]
        private class CVirtualDesktopManager
        {
        }
    }
}
