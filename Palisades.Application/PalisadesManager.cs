using Palisades.Helpers;
using Palisades.Model;
using Palisades.View;
using Palisades.ViewModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Interop;
using System.Xml.Serialization;

namespace Palisades
{
    internal static class PalisadesManager
    {
        public static readonly Dictionary<string, Palisade> palisades = new();

        public static void LoadPalisades()
        {
            string saveDirectory = PDirectory.GetPalisadesDirectory();
            PDirectory.EnsureExists(saveDirectory);

            List<(string DirectoryPath, PalisadeModel Model)> loadedEntries = new();

            foreach (string palisadeDirectory in Directory.GetDirectories(saveDirectory))
            {
                string statePath = Path.Combine(palisadeDirectory, "state.xml");
                if (!File.Exists(statePath))
                {
                    continue;
                }

                try
                {
                    XmlSerializer deserializer = new(typeof(PalisadeModel));
                    using StreamReader reader = new(statePath);
                    if (deserializer.Deserialize(reader) is PalisadeModel model)
                    {
                        loadedEntries.Add((palisadeDirectory, model));
                    }
                }
                catch
                {
                    // Ignore invalid saved states and continue loading the rest.
                }
            }

            List<(string DirectoryPath, PalisadeModel Model)> placeholderEntries = loadedEntries
                .Where(entry => IsPlaceholderFence(entry.Model))
                .ToList();

            if (placeholderEntries.Count > 0)
            {
                if (loadedEntries.Count > placeholderEntries.Count)
                {
                    foreach ((string DirectoryPath, PalisadeModel _) in placeholderEntries)
                    {
                        DeleteSavedFenceDirectory(DirectoryPath);
                    }

                    loadedEntries = loadedEntries
                        .Where(entry => !IsPlaceholderFence(entry.Model))
                        .ToList();
                }
                else if (placeholderEntries.Count > 1)
                {
                    foreach ((string DirectoryPath, PalisadeModel _) in placeholderEntries.Skip(1))
                    {
                        DeleteSavedFenceDirectory(DirectoryPath);
                    }

                    loadedEntries = placeholderEntries.Take(1).ToList();
                }
            }

            foreach ((string _, PalisadeModel loadedModel) in loadedEntries)
            {
                if (palisades.ContainsKey(loadedModel.Identifier))
                {
                    continue;
                }

                palisades.Add(loadedModel.Identifier, new Palisade(new PalisadeViewModel(loadedModel)));
            }
        }

        private static bool IsPlaceholderFence(PalisadeModel model)
        {
            return string.Equals(model.Name?.Trim(), "No name", StringComparison.OrdinalIgnoreCase)
                && (model.Shortcuts == null || model.Shortcuts.Count == 0);
        }

        private static void DeleteSavedFenceDirectory(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                return;
            }

            try
            {
                Directory.Delete(directoryPath, true);
            }
            catch
            {
                // Best-effort cleanup only.
            }
        }

        private static void LoadPalisade(PalisadeViewModel initialModel)
        {
            Palisade palisade = new(initialModel);
            palisades.Add(initialModel.Identifier, palisade);
        }

        public static void CreatePalisade()
        {
            PalisadeViewModel viewModel = new();
            palisades.Add(viewModel.Identifier, new Palisade(viewModel));
            viewModel.Save();
            ApplyDesktopVisibilityForCurrentDesktop();
        }

        public static void ApplyDesktopVisibilityForCurrentDesktop()
        {
            string currentDesktopId = VirtualDesktopHelper.GetCurrentDesktopIdString();
            foreach (KeyValuePair<string, Palisade> entry in palisades.ToList())
            {
                if (entry.Value.DataContext is PalisadeViewModel viewModel)
                {
                    viewModel.ApplyDesktopVisibility(entry.Value, currentDesktopId);
                }
            }
        }

        public static void HidePalisade(string identifier)
        {
            palisades.TryGetValue(identifier, out Palisade? palisade);
            if (palisade == null)
            {
                return;
            }

            if (palisade.DataContext is PalisadeViewModel viewModel)
            {
                viewModel.SetHiddenByUser(true);
            }

            palisade.Hide();
        }

        public static void ShowPalisade(string identifier)
        {
            palisades.TryGetValue(identifier, out Palisade? palisade);
            if (palisade == null)
            {
                return;
            }

            if (palisade.DataContext is PalisadeViewModel viewModel)
            {
                viewModel.SetHiddenByUser(false);
            }

            if (!palisade.IsVisible)
            {
                palisade.Show();
            }

            if (palisade.WindowState == System.Windows.WindowState.Minimized)
            {
                palisade.WindowState = System.Windows.WindowState.Normal;
            }

            palisade.Activate();
        }

        public static void MovePalisadeToDesktop(string identifier, string desktopId)
        {
            if (!Guid.TryParse(desktopId, out Guid targetDesktopId))
            {
                return;
            }

            palisades.TryGetValue(identifier, out Palisade? palisade);
            if (palisade == null)
            {
                return;
            }

            IntPtr windowHandle = new WindowInteropHelper(palisade).Handle;
            if (windowHandle == IntPtr.Zero)
            {
                return;
            }

            VirtualDesktopHelper.TryMoveWindowToDesktop(windowHandle, targetDesktopId);
        }

        public static void MovePalisadeToPreviousDesktop(string identifier)
        {
            MovePalisadeToAdjacentDesktop(identifier, moveToNextDesktop: false);
        }

        public static void MovePalisadeToNextDesktop(string identifier)
        {
            MovePalisadeToAdjacentDesktop(identifier, moveToNextDesktop: true);
        }

        private static void MovePalisadeToAdjacentDesktop(string identifier, bool moveToNextDesktop)
        {
            palisades.TryGetValue(identifier, out Palisade? palisade);
            if (palisade == null)
            {
                return;
            }

            IntPtr windowHandle = new WindowInteropHelper(palisade).Handle;
            if (windowHandle == IntPtr.Zero)
            {
                return;
            }

            if (moveToNextDesktop)
            {
                VirtualDesktopHelper.TryMoveWindowToNextDesktop(windowHandle);
            }
            else
            {
                VirtualDesktopHelper.TryMoveWindowToPreviousDesktop(windowHandle);
            }
        }

        public static void DeletePalisade(string identifier)
        {
            palisades.TryGetValue(identifier, out Palisade? palisade);
            if (palisade == null)
            {
                return;
            }
            if (palisade.DataContext != null)
            {
                ((PalisadeViewModel)palisade.DataContext).Delete();
            }

            palisade.Close();
            palisades.Remove(identifier);

        }

        public static Palisade GetPalisade(string identifier)
        {
            palisades.TryGetValue(identifier, out Palisade? palisade);
            if (palisade == null)
            {
                throw new KeyNotFoundException(identifier);
            }
            return palisade;
        }
    }
}
