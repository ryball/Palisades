using Palisades.Helpers;
using Palisades.Model;
using Palisades.View;
using Palisades.ViewModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
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

            ApplyTabbedVisibility();
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

                if (GetViewModelsByGroup(viewModel.TabGroupId).Count > 1)
                {
                    ActivateTabbedFence(identifier);
                    return;
                }
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

        public static IReadOnlyList<PalisadeTabInfo> GetJoinedTabsFor(string identifier)
        {
            PalisadeViewModel? currentViewModel = TryGetViewModel(identifier);
            if (currentViewModel == null)
            {
                return Array.Empty<PalisadeTabInfo>();
            }

            string groupId = currentViewModel.TabGroupId;

            return GetViewModelsByGroup(groupId)
                .Select(viewModel => new PalisadeTabInfo
                {
                    Identifier = viewModel.Identifier,
                    Name = viewModel.Name,
                    IsCurrent = string.Equals(viewModel.Identifier, identifier, StringComparison.OrdinalIgnoreCase)
                })
                .ToList();
        }

        public static IReadOnlyList<PalisadeTabInfo> GetJoinTargetsFor(string identifier)
        {
            PalisadeViewModel? currentViewModel = TryGetViewModel(identifier);
            if (currentViewModel == null)
            {
                return Array.Empty<PalisadeTabInfo>();
            }

            return palisades.Values
                .Select(window => window.DataContext as PalisadeViewModel)
                .Where(viewModel => viewModel != null
                    && !string.Equals(viewModel.Identifier, identifier, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(viewModel.TabGroupId, currentViewModel.TabGroupId, StringComparison.OrdinalIgnoreCase))
                .OrderBy(viewModel => viewModel!.Name, StringComparer.CurrentCultureIgnoreCase)
                .Select(viewModel => new PalisadeTabInfo
                {
                    Identifier = viewModel!.Identifier,
                    Name = viewModel.Name,
                    IsCurrent = false
                })
                .ToList();
        }

        public static void JoinPalisadesAsTabs(string hostIdentifier, string targetIdentifier)
        {
            PalisadeViewModel? hostViewModel = TryGetViewModel(hostIdentifier);
            PalisadeViewModel? targetViewModel = TryGetViewModel(targetIdentifier);
            if (hostViewModel == null || targetViewModel == null)
            {
                return;
            }

            string hostGroupId = hostViewModel.TabGroupId;
            string targetGroupId = targetViewModel.TabGroupId;
            if (string.Equals(hostGroupId, targetGroupId, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            List<PalisadeViewModel> hostGroupMembers = GetViewModelsByGroup(hostGroupId);
            List<PalisadeViewModel> targetGroupMembers = GetViewModelsByGroup(targetGroupId);
            List<PalisadeViewModel> affectedViewModels = hostGroupMembers
                .Concat(targetGroupMembers)
                .GroupBy(viewModel => viewModel.Identifier, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();

            int nextTabOrder = hostGroupMembers.Count;
            foreach (PalisadeViewModel groupMember in targetGroupMembers)
            {
                groupMember.TabGroupId = hostGroupId;
                groupMember.TabOrder = nextTabOrder++;
                groupMember.ActiveTabIdentifier = hostIdentifier;
                SyncWindowLayout(hostViewModel, groupMember);
                groupMember.Save();
            }

            foreach (PalisadeViewModel groupMember in hostGroupMembers)
            {
                groupMember.ActiveTabIdentifier = hostIdentifier;
                SyncWindowLayout(hostViewModel, groupMember);
                groupMember.Save();
            }

            ActivateTabbedFence(hostIdentifier);
            RefreshTabbedStateFor(affectedViewModels.Select(viewModel => viewModel.Identifier));
        }

        public static void SplitPalisadeFromTabs(string identifier, Point? detachedScreenPosition = null)
        {
            PalisadeViewModel? viewModel = TryGetViewModel(identifier);
            if (viewModel == null)
            {
                return;
            }

            string previousGroupId = viewModel.TabGroupId;
            List<PalisadeViewModel> groupMembers = GetViewModelsByGroup(previousGroupId);
            if (groupMembers.Count <= 1)
            {
                return;
            }

            PalisadeViewModel? remainingViewModel = groupMembers.FirstOrDefault(member => !string.Equals(member.Identifier, identifier, StringComparison.OrdinalIgnoreCase));
            viewModel.TabGroupId = viewModel.Identifier;
            viewModel.TabOrder = 0;
            viewModel.ActiveTabIdentifier = viewModel.Identifier;

            if (detachedScreenPosition.HasValue)
            {
                Point targetPosition = detachedScreenPosition.Value;
                double horizontalOffset = Math.Min(viewModel.Width, 240) / 2d;
                double verticalOffset = Math.Max(viewModel.HeaderHeight, 40) / 2d;
                viewModel.FenceX = Math.Max(0, (int)Math.Round(targetPosition.X - horizontalOffset));
                viewModel.FenceY = Math.Max(0, (int)Math.Round(targetPosition.Y - verticalOffset));
            }
            else
            {
                viewModel.FenceX += 40;
                viewModel.FenceY += 40;
            }

            viewModel.SetHiddenByUser(false);
            viewModel.Save();

            int remainingOrder = 0;
            foreach (PalisadeViewModel groupMember in groupMembers.Where(member => !string.Equals(member.Identifier, identifier, StringComparison.OrdinalIgnoreCase)))
            {
                groupMember.TabOrder = remainingOrder++;
                if (string.Equals(groupMember.ActiveTabIdentifier, identifier, StringComparison.OrdinalIgnoreCase))
                {
                    groupMember.ActiveTabIdentifier = remainingViewModel?.Identifier ?? groupMember.Identifier;
                }

                groupMember.Save();
            }

            if (remainingViewModel != null)
            {
                ActivateTabbedFence(remainingViewModel.Identifier);
            }

            ShowPalisade(identifier);
            RefreshTabbedStateFor(groupMembers.Select(member => member.Identifier).Append(identifier));
        }

        public static void ReorderTabbedFence(string draggedIdentifier, string targetIdentifier, bool placeAfter)
        {
            if (string.IsNullOrWhiteSpace(draggedIdentifier)
                || string.IsNullOrWhiteSpace(targetIdentifier)
                || string.Equals(draggedIdentifier, targetIdentifier, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            PalisadeViewModel? draggedViewModel = TryGetViewModel(draggedIdentifier);
            PalisadeViewModel? targetViewModel = TryGetViewModel(targetIdentifier);
            if (draggedViewModel == null || targetViewModel == null)
            {
                return;
            }

            if (!string.Equals(draggedViewModel.TabGroupId, targetViewModel.TabGroupId, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            List<PalisadeViewModel> groupMembers = GetViewModelsByGroup(draggedViewModel.TabGroupId);
            if (groupMembers.Count <= 1)
            {
                return;
            }

            PalisadeViewModel? draggedMember = groupMembers.FirstOrDefault(member => string.Equals(member.Identifier, draggedIdentifier, StringComparison.OrdinalIgnoreCase));
            PalisadeViewModel? targetMember = groupMembers.FirstOrDefault(member => string.Equals(member.Identifier, targetIdentifier, StringComparison.OrdinalIgnoreCase));
            if (draggedMember == null || targetMember == null)
            {
                return;
            }

            groupMembers.Remove(draggedMember);
            int targetIndex = groupMembers.FindIndex(member => string.Equals(member.Identifier, targetIdentifier, StringComparison.OrdinalIgnoreCase));
            if (targetIndex < 0)
            {
                return;
            }

            int insertIndex = placeAfter ? targetIndex + 1 : targetIndex;
            insertIndex = Math.Clamp(insertIndex, 0, groupMembers.Count);
            groupMembers.Insert(insertIndex, draggedMember);

            for (int index = 0; index < groupMembers.Count; index++)
            {
                groupMembers[index].TabOrder = index;
                groupMembers[index].Save();
            }

            ApplyTabbedVisibility(draggedViewModel.TabGroupId);
            RefreshTabbedStateFor(groupMembers.Select(member => member.Identifier));
        }

        public static void ActivateTabbedFence(string identifier)
        {
            PalisadeViewModel? targetViewModel = TryGetViewModel(identifier);
            if (targetViewModel == null)
            {
                return;
            }

            List<PalisadeViewModel> groupMembers = GetViewModelsByGroup(targetViewModel.TabGroupId);
            if (groupMembers.Count <= 1)
            {
                ShowPalisade(identifier);
                return;
            }

            PalisadeViewModel? currentVisibleViewModel = groupMembers.FirstOrDefault(member => palisades.TryGetValue(member.Identifier, out Palisade? candidateWindow) && candidateWindow.IsVisible);
            if (currentVisibleViewModel != null)
            {
                SyncWindowLayout(currentVisibleViewModel, targetViewModel);
            }

            foreach (PalisadeViewModel groupMember in groupMembers)
            {
                groupMember.ActiveTabIdentifier = identifier;
                groupMember.Save();
            }

            ApplyTabbedVisibility(targetViewModel.TabGroupId);

            if (palisades.TryGetValue(identifier, out Palisade? targetWindow))
            {
                if (targetWindow.WindowState == System.Windows.WindowState.Minimized)
                {
                    targetWindow.WindowState = System.Windows.WindowState.Normal;
                }

                targetWindow.Activate();
            }

            RefreshTabbedStateFor(groupMembers.Select(member => member.Identifier));
        }

        public static void ApplyTabbedVisibility(string? specificGroupId = null)
        {
            string currentDesktopId = VirtualDesktopHelper.GetCurrentDesktopIdString();
            IEnumerable<IGrouping<string, PalisadeViewModel>> groupedViewModels = palisades.Values
                .Select(window => window.DataContext as PalisadeViewModel)
                .Where(viewModel => viewModel != null)
                .GroupBy(viewModel => viewModel!.TabGroupId, StringComparer.OrdinalIgnoreCase);

            if (!string.IsNullOrWhiteSpace(specificGroupId))
            {
                groupedViewModels = groupedViewModels.Where(group => string.Equals(group.Key, specificGroupId, StringComparison.OrdinalIgnoreCase));
            }

            foreach (IGrouping<string, PalisadeViewModel> group in groupedViewModels)
            {
                List<PalisadeViewModel> members = group.ToList();
                if (members.Count <= 1)
                {
                    continue;
                }

                string activeIdentifier = ResolveActiveTabIdentifier(group.Key, members);
                foreach (PalisadeViewModel member in members)
                {
                    if (!palisades.TryGetValue(member.Identifier, out Palisade? window))
                    {
                        continue;
                    }

                    if (string.Equals(member.Identifier, activeIdentifier, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!member.IsHiddenByUser && member.IsVisibleOnDesktop(currentDesktopId) && !window.IsVisible)
                        {
                            window.Show();
                        }
                    }
                    else
                    {
                        window.Hide();
                    }
                }
            }
        }

        private static PalisadeViewModel? TryGetViewModel(string identifier)
        {
            return palisades.TryGetValue(identifier, out Palisade? palisade) ? palisade.DataContext as PalisadeViewModel : null;
        }

        private static List<PalisadeViewModel> GetViewModelsByGroup(string? groupId)
        {
            if (string.IsNullOrWhiteSpace(groupId))
            {
                return new List<PalisadeViewModel>();
            }

            List<PalisadeViewModel> members = palisades.Values
                .Select(window => window.DataContext as PalisadeViewModel)
                .Where(viewModel => viewModel != null && string.Equals(viewModel.TabGroupId, groupId, StringComparison.OrdinalIgnoreCase))
                .Cast<PalisadeViewModel>()
                .ToList();

            List<PalisadeViewModel> orderedMembers = members
                .Select((viewModel, index) => new { viewModel, index })
                .OrderBy(item => item.viewModel.TabOrder)
                .ThenBy(item => item.index)
                .Select(item => item.viewModel)
                .ToList();

            bool changed = false;
            for (int index = 0; index < orderedMembers.Count; index++)
            {
                if (orderedMembers[index].TabOrder == index)
                {
                    continue;
                }

                orderedMembers[index].TabOrder = index;
                changed = true;
            }

            if (changed)
            {
                foreach (PalisadeViewModel member in orderedMembers)
                {
                    member.Save();
                }
            }

            return orderedMembers;
        }

        private static void RefreshTabbedStateFor(IEnumerable<string> identifiers)
        {
            foreach (string identifier in identifiers.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                TryGetViewModel(identifier)?.RefreshTabbedState();
            }
        }

        private static string ResolveActiveTabIdentifier(string groupId)
        {
            return ResolveActiveTabIdentifier(groupId, GetViewModelsByGroup(groupId));
        }

        private static string ResolveActiveTabIdentifier(string groupId, IReadOnlyCollection<PalisadeViewModel> members)
        {
            PalisadeViewModel? activeViewModel = members.FirstOrDefault(member => members.Any(candidate => string.Equals(candidate.Identifier, member.ActiveTabIdentifier, StringComparison.OrdinalIgnoreCase)));
            if (activeViewModel != null)
            {
                return activeViewModel.ActiveTabIdentifier;
            }

            PalisadeViewModel? groupLeader = members.FirstOrDefault(member => string.Equals(member.Identifier, groupId, StringComparison.OrdinalIgnoreCase));
            if (groupLeader != null)
            {
                return groupLeader.Identifier;
            }

            return members.FirstOrDefault()?.Identifier ?? string.Empty;
        }

        private static void SyncWindowLayout(PalisadeViewModel sourceViewModel, PalisadeViewModel targetViewModel)
        {
            if (sourceViewModel == null || targetViewModel == null || string.Equals(sourceViewModel.Identifier, targetViewModel.Identifier, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            targetViewModel.FenceX = sourceViewModel.FenceX;
            targetViewModel.FenceY = sourceViewModel.FenceY;
            targetViewModel.Width = sourceViewModel.Width;
            targetViewModel.Height = sourceViewModel.Height;
            targetViewModel.IsCollapsed = sourceViewModel.IsCollapsed;
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

        public static bool TryJoinPalisadeByOverlap(string draggedIdentifier)
        {
            if (!palisades.TryGetValue(draggedIdentifier, out Palisade? draggedWindow) || draggedWindow == null)
            {
                return false;
            }

            PalisadeViewModel? draggedViewModel = draggedWindow.DataContext as PalisadeViewModel;
            if (draggedViewModel == null)
            {
                return false;
            }

            Rect draggedBounds = new(draggedWindow.Left, draggedWindow.Top, Math.Max(draggedWindow.Width, draggedWindow.ActualWidth), Math.Max(draggedWindow.Height, draggedWindow.ActualHeight));
            Point draggedCenter = new(draggedBounds.Left + (draggedBounds.Width / 2d), draggedBounds.Top + (draggedBounds.Height / 2d));

            string? bestTargetIdentifier = null;
            double bestOverlapScore = 0d;

            foreach (KeyValuePair<string, Palisade> entry in palisades)
            {
                if (string.Equals(entry.Key, draggedIdentifier, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Palisade otherWindow = entry.Value;
                if (!otherWindow.IsVisible || otherWindow.DataContext is not PalisadeViewModel otherViewModel)
                {
                    continue;
                }

                if (string.Equals(otherViewModel.TabGroupId, draggedViewModel.TabGroupId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Rect otherBounds = new(otherWindow.Left, otherWindow.Top, Math.Max(otherWindow.Width, otherWindow.ActualWidth), Math.Max(otherWindow.Height, otherWindow.ActualHeight));
                Rect intersection = Rect.Intersect(draggedBounds, otherBounds);
                if (intersection.IsEmpty || intersection.Width <= 0 || intersection.Height <= 0)
                {
                    continue;
                }

                double overlapArea = intersection.Width * intersection.Height;
                double minimumArea = Math.Max(2500d, Math.Min((draggedBounds.Width * draggedBounds.Height) * 0.1d, 30000d));
                bool centerInsideTarget = otherBounds.Contains(draggedCenter);
                if (!centerInsideTarget && overlapArea < minimumArea)
                {
                    continue;
                }

                double overlapScore = centerInsideTarget ? overlapArea + 1000000d : overlapArea;
                if (overlapScore > bestOverlapScore)
                {
                    bestOverlapScore = overlapScore;
                    bestTargetIdentifier = entry.Key;
                }
            }

            if (string.IsNullOrWhiteSpace(bestTargetIdentifier))
            {
                return false;
            }

            JoinPalisadesAsTabs(bestTargetIdentifier, draggedIdentifier);
            ActivateTabbedFence(draggedIdentifier);
            return true;
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

            string previousGroupId = (palisade.DataContext as PalisadeViewModel)?.TabGroupId ?? identifier;
            if (palisade.DataContext != null)
            {
                ((PalisadeViewModel)palisade.DataContext).Delete();
            }

            palisade.Close();
            palisades.Remove(identifier);

            List<PalisadeViewModel> remainingGroupMembers = GetViewModelsByGroup(previousGroupId);
            if (remainingGroupMembers.Count > 0)
            {
                string remainingIdentifier = ResolveActiveTabIdentifier(previousGroupId, remainingGroupMembers);
                ActivateTabbedFence(remainingIdentifier);
            }

            RefreshTabbedStateFor(remainingGroupMembers.Select(viewModel => viewModel.Identifier));
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
