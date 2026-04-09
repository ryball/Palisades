using GongSolutions.Wpf.DragDrop;
using Microsoft.VisualBasic;
using Palisades.Helpers;
using Palisades.Model;
using Palisades.View;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Xml.Serialization;

namespace Palisades.ViewModel
{
    public class PalisadeViewModel : INotifyPropertyChanged, IDropTarget, IDragSource
    {
        #region Attributs
        private const string ShortcutDragDataFormat = "Palisades.ShortcutDrag";

        private readonly DefaultDropHandler defaultDropHandler = new();
        private readonly DefaultDragHandler defaultDragHandler = new();
        private readonly PalisadeModel model;

        private ICollectionView groupedShortcuts = null!;
        private volatile bool shouldSave;
        private Shortcut? selectedShortcut;
        private Shortcut? lastRemovedShortcut;
        private int lastRemovedShortcutIndex = -1;
        private bool isHiddenByUser;
        #endregion

        #region Accessors
        public string Identifier
        {
            get { return model.Identifier; }
            set { model.Identifier = value; OnPropertyChanged(); Save(); }
        }

        public string Name
        {
            get { return model.Name; }
            set
            {
                model.Name = value;
                OnPropertyChanged();
                RefreshTabbedState();
                Save();
            }
        }

        public string TabGroupId
        {
            get { return model.TabGroupId; }
            set
            {
                string normalized = string.IsNullOrWhiteSpace(value) ? Identifier : value.Trim();
                if (string.Equals(model.TabGroupId, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                model.TabGroupId = normalized;
                RefreshTabbedState();
                Save();
            }
        }

        public string ActiveTabIdentifier
        {
            get { return model.ActiveTabIdentifier; }
            set
            {
                string normalized = string.IsNullOrWhiteSpace(value) ? Identifier : value.Trim();
                if (string.Equals(model.ActiveTabIdentifier, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                model.ActiveTabIdentifier = normalized;
                RefreshTabbedState();
                Save();
            }
        }

        public int TabOrder
        {
            get { return model.TabOrder; }
            set
            {
                int normalized = Math.Max(0, value);
                if (model.TabOrder == normalized)
                {
                    return;
                }

                model.TabOrder = normalized;
                RefreshTabbedState();
                Save();
            }
        }

        public int FenceX
        {
            get { return model.FenceX; }
            set { model.FenceX = value; OnPropertyChanged(); Save(); }
        }

        public int FenceY
        {
            get { return model.FenceY; }
            set { model.FenceY = value; OnPropertyChanged(); Save(); }
        }

        public int Width
        {
            get { return model.Width; }
            set { model.Width = value; OnPropertyChanged(); Save(); }
        }

        public int Height
        {
            get { return IsCollapsed ? CollapsedHeight : model.Height; }
            set
            {
                int normalized = Math.Max(value, CollapsedHeight);
                if (IsCollapsed)
                {
                    OnPropertyChanged();
                    return;
                }

                if (model.Height == normalized)
                {
                    return;
                }

                model.Height = normalized;
                OnPropertyChanged();
                Save();
            }
        }

        public int HeaderHeight
        {
            get { return model.HeaderHeight; }
            set
            {
                int normalized = Math.Clamp(value, PalisadeModel.MinHeaderHeight, PalisadeModel.MaxHeaderHeight);
                if (model.HeaderHeight == normalized)
                {
                    return;
                }

                model.HeaderHeight = normalized;
                if (model.Height < CollapsedHeight)
                {
                    model.Height = CollapsedHeight;
                }

                OnPropertyChanged();
                OnPropertyChanged(nameof(Height));
                OnPropertyChanged(nameof(TitleFontSize));
                Save();
            }
        }

        public double TitleFontSize
        {
            get { return Math.Clamp(HeaderHeight * 0.53d, 18d, 48d); }
        }

        public bool IsCollapsed
        {
            get { return model.IsCollapsed; }
            set
            {
                if (model.IsCollapsed == value)
                {
                    return;
                }

                if (!value && model.Height < CollapsedHeight)
                {
                    model.Height = CollapsedHeight;
                }

                model.IsCollapsed = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(Height));
                OnPropertyChanged(nameof(BodyVisibility));
                OnPropertyChanged(nameof(CollapseMenuHeader));
                OnPropertyChanged(nameof(CollapseButtonText));
                OnPropertyChanged(nameof(WindowResizeMode));
                Save();
            }
        }

        public Visibility BodyVisibility
        {
            get { return IsCollapsed ? Visibility.Collapsed : Visibility.Visible; }
        }

        public string CollapseMenuHeader
        {
            get { return IsCollapsed ? "Expand fence" : "Collapse to title bar"; }
        }

        public string CollapseButtonText
        {
            get { return IsCollapsed ? "▾" : "▴"; }
        }

        public bool SupportsMultipleDesktops
        {
            get { return VirtualDesktopHelper.HasMultipleDesktops(); }
        }

        public string StartupSettingLabel
        {
            get { return $"Start {AppBranding.DisplayName} when I sign in to Windows"; }
        }

        public bool StartWithWindows
        {
            get { return StartupLaunchHelper.IsEnabled(); }
            set
            {
                bool currentValue = StartupLaunchHelper.IsEnabled();
                if (currentValue == value)
                {
                    return;
                }

                StartupLaunchHelper.SetEnabled(value);
                OnPropertyChanged();
            }
        }

        public IEnumerable<VirtualDesktopInfo> AvailableDesktopTargets
        {
            get
            {
                return VirtualDesktopHelper.GetDesktops()
                    .Select(desktop => new VirtualDesktopInfo
                    {
                        DesktopId = desktop.DesktopId,
                        Name = desktop.Name,
                        IsCurrent = desktop.IsCurrent,
                        IsSelected = IsVisibleOnDesktop(desktop.DesktopId)
                    })
                    .ToList();
            }
        }

        public IEnumerable<PalisadeTabInfo> JoinedTabs
        {
            get { return PalisadesManager.GetJoinedTabsFor(Identifier); }
        }

        public IEnumerable<PalisadeTabInfo> AvailableJoinTargets
        {
            get { return PalisadesManager.GetJoinTargetsFor(Identifier); }
        }

        public bool HasTabbedGroup
        {
            get { return JoinedTabs.Skip(1).Any(); }
        }

        public Visibility TabStripVisibility
        {
            get { return HasTabbedGroup ? Visibility.Visible : Visibility.Collapsed; }
        }

        public Visibility TitleVisibility
        {
            get { return HasTabbedGroup ? Visibility.Collapsed : Visibility.Visible; }
        }

        public ResizeMode WindowResizeMode
        {
            get { return IsCollapsed ? ResizeMode.NoResize : ResizeMode.CanResize; }
        }

        private int CollapsedHeight
        {
            get { return Math.Max(HeaderHeight, PalisadeModel.MinHeaderHeight); }
        }

        public Color HeaderColor
        {
            get { return model.HeaderColor; }
            set { model.HeaderColor = value; OnPropertyChanged(); Save(); }
        }

        public Color BodyColor
        {
            get { return model.BodyColor; }
            set { model.BodyColor = value; OnPropertyChanged(); Save(); }
        }

        public SolidColorBrush TitleColor
        {
            get => new(model.TitleColor);
            set { model.TitleColor = value.Color; OnPropertyChanged(); Save(); }
        }
        public SolidColorBrush LabelsColor
        {
            get => new(model.LabelsColor);
            set { model.LabelsColor = value.Color; OnPropertyChanged(); Save(); }
        }

        public ObservableCollection<Shortcut> Shortcuts
        {
            get { return model.Shortcuts; }
            set
            {
                if (ReferenceEquals(model.Shortcuts, value))
                {
                    return;
                }

                UnsubscribeFromShortcuts(model.Shortcuts);
                model.Shortcuts = value;
                SubscribeToShortcuts(model.Shortcuts);
                ConfigureGroupedShortcuts();
                OnPropertyChanged();
                OnPropertyChanged(nameof(GroupedShortcuts));
                OnPropertyChanged(nameof(AvailableGroups));
                Save();
            }
        }

        public ICollectionView GroupedShortcuts
        {
            get { return groupedShortcuts; }
        }

        public IEnumerable<string> AvailableGroups
        {
            get
            {
                return model.GroupStates
                    .Select(state => NormalizeGroupName(state.GroupName))
                    .Concat(Shortcuts.Select(shortcut => NormalizeGroupName(shortcut.GroupName)))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(groupName => groupName)
                    .ToList();
            }
        }

        public IEnumerable<string> AvailableTypes
        {
            get
            {
                return model.Types
                    .Concat(Shortcuts.Select(shortcut => NormalizeTypeName(shortcut.TypeName)))
                    .Where(typeName => !string.IsNullOrWhiteSpace(typeName))
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .OrderBy(typeName => typeName, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
            }
        }

        public Shortcut? SelectedShortcut
        {
            get => selectedShortcut;
            set
            {
                selectedShortcut = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasSelectedShortcut));
                OnPropertyChanged(nameof(SelectedShortcutGroupName));
                CommandManager.InvalidateRequerySuggested();
                Save();
            }
        }

        public bool HasSelectedShortcut
        {
            get { return SelectedShortcut != null; }
        }

        public bool CanUndoShortcutRemoval
        {
            get { return lastRemovedShortcut != null && lastRemovedShortcutIndex >= 0; }
        }

        public string SelectedShortcutGroupName
        {
            get { return SelectedShortcut?.GroupName ?? string.Empty; }
            set
            {
                if (SelectedShortcut == null)
                {
                    return;
                }

                string normalized = NormalizeGroupName(value);
                if (string.Equals(SelectedShortcut.GroupName, normalized, StringComparison.Ordinal))
                {
                    return;
                }

                SelectedShortcut.GroupName = normalized;
                RegisterGroup(normalized);
                RefreshGroups();
                OnPropertyChanged();
                OnPropertyChanged(nameof(AvailableGroups));
                Save();
            }
        }
        #endregion

        public PalisadeViewModel() : this(new PalisadeModel()) { }

        public PalisadeViewModel(PalisadeModel model)
        {
            this.model = model;
            this.model.EnsureDefaults();
            EnsureCurrentDesktopAssignment();

            SubscribeToShortcuts(this.model.Shortcuts);
            ConfigureGroupedShortcuts();

            OnPropertyChanged();
            OnPropertyChanged(nameof(GroupedShortcuts));
            OnPropertyChanged(nameof(AvailableGroups));
            OnPropertyChanged(nameof(AvailableTypes));
            OnPropertyChanged(nameof(AvailableDesktopTargets));
            RefreshTabbedState();
            Save();

            Thread saveThread = new(SaveAsync)
            {
                IsBackground = true,
                Name = $"PalisadeSave-{Identifier}"
            };
            saveThread.Start();
        }

        #region Methods
        public void Save()
        {
            shouldSave = true;
        }

        public void RefreshTabbedState()
        {
            OnPropertyChanged(nameof(JoinedTabs));
            OnPropertyChanged(nameof(AvailableJoinTargets));
            OnPropertyChanged(nameof(HasTabbedGroup));
            OnPropertyChanged(nameof(TabStripVisibility));
            OnPropertyChanged(nameof(TitleVisibility));
        }

        public bool IsHiddenByUser
        {
            get { return isHiddenByUser; }
        }

        public void SetHiddenByUser(bool isHidden)
        {
            isHiddenByUser = isHidden;
        }

        public bool IsVisibleOnDesktop(string? desktopId)
        {
            string normalized = NormalizeDesktopId(desktopId);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return true;
            }

            if (model.VisibleDesktopIds.Count == 0)
            {
                return true;
            }

            return model.VisibleDesktopIds.Any(id => string.Equals(id, normalized, StringComparison.OrdinalIgnoreCase));
        }

        public void ApplyDesktopVisibility(View.Palisade palisade, string? currentDesktopId)
        {
            if (!SupportsMultipleDesktops)
            {
                return;
            }

            string normalizedDesktopId = NormalizeDesktopId(currentDesktopId);
            if (string.IsNullOrWhiteSpace(normalizedDesktopId))
            {
                return;
            }

            bool shouldBeVisible = IsVisibleOnDesktop(normalizedDesktopId);
            IntPtr windowHandle = new WindowInteropHelper(palisade).Handle;

            if (shouldBeVisible)
            {
                if (windowHandle != IntPtr.Zero && Guid.TryParse(normalizedDesktopId, out Guid currentDesktopGuid))
                {
                    VirtualDesktopHelper.TryMoveWindowToDesktop(windowHandle, currentDesktopGuid);
                }

                if (!isHiddenByUser)
                {
                    if (!palisade.IsVisible)
                    {
                        palisade.Show();
                    }

                    if (palisade.WindowState == WindowState.Minimized)
                    {
                        palisade.WindowState = WindowState.Normal;
                    }
                }
            }
            else if (palisade.IsVisible)
            {
                palisade.Hide();
            }
        }

        public void ToggleDesktopVisibility(string? desktopId)
        {
            string normalized = NormalizeDesktopId(desktopId);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            string? existing = model.VisibleDesktopIds.FirstOrDefault(id => string.Equals(id, normalized, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                if (model.VisibleDesktopIds.Count == 1)
                {
                    return;
                }

                model.VisibleDesktopIds.Remove(existing);
            }
            else
            {
                model.VisibleDesktopIds.Add(normalized);
            }

            OnPropertyChanged(nameof(AvailableDesktopTargets));
            Save();
            PalisadesManager.ApplyDesktopVisibilityForCurrentDesktop();
        }

        private void EnsureCurrentDesktopAssignment()
        {
            if (!SupportsMultipleDesktops || model.VisibleDesktopIds.Count > 0)
            {
                return;
            }

            string currentDesktopId = NormalizeDesktopId(VirtualDesktopHelper.GetCurrentDesktopIdString());
            if (string.IsNullOrWhiteSpace(currentDesktopId))
            {
                return;
            }

            model.VisibleDesktopIds.Add(currentDesktopId);
        }

        public void Delete()
        {
            string saveDirectory = PDirectory.GetPalisadeDirectory(Identifier);
            Directory.Delete(Path.Combine(saveDirectory), true);
        }

        public bool IsGroupExpanded(string? groupName)
        {
            return model.EnsureGroupState(NormalizeGroupName(groupName)).IsExpanded;
        }

        public void SetGroupExpanded(string? groupName, bool isExpanded)
        {
            PalisadeGroupState state = model.EnsureGroupState(NormalizeGroupName(groupName));
            if (state.IsExpanded == isExpanded)
            {
                return;
            }

            state.IsExpanded = isExpanded;
            Save();
        }

        public void MoveSelectedShortcutToGroup(string? groupName)
        {
            if (SelectedShortcut == null)
            {
                return;
            }

            SelectedShortcutGroupName = NormalizeGroupName(groupName);
        }

        public void MoveSelectedShortcutToType(string? typeName)
        {
            if (SelectedShortcut == null)
            {
                return;
            }

            string normalized = NormalizeTypeName(typeName);
            SelectedShortcut.TypeName = normalized;
            RegisterType(normalized);
            OnPropertyChanged(nameof(AvailableTypes));
            Save();
        }

        private void ToggleCollapsedState()
        {
            IsCollapsed = !IsCollapsed;
        }

        private void SortShortcutsByName(bool descending)
        {
            List<Shortcut> orderedShortcuts = descending
                ? Shortcuts.OrderByDescending(shortcut => shortcut.Name, StringComparer.CurrentCultureIgnoreCase).ToList()
                : Shortcuts.OrderBy(shortcut => shortcut.Name, StringComparer.CurrentCultureIgnoreCase).ToList();

            ApplyShortcutOrder(orderedShortcuts);
        }

        private void SortShortcutsByType()
        {
            List<Shortcut> orderedShortcuts = Shortcuts
                .OrderBy(shortcut => GetTypeSortKey(shortcut.TypeName), StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(shortcut => shortcut.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            ApplyShortcutOrder(orderedShortcuts);
        }

        private void ApplyShortcutOrder(IList<Shortcut> orderedShortcuts)
        {
            if (orderedShortcuts.Count != Shortcuts.Count)
            {
                return;
            }

            Shortcut? currentSelection = SelectedShortcut;
            for (int targetIndex = 0; targetIndex < orderedShortcuts.Count; targetIndex++)
            {
                Shortcut shortcut = orderedShortcuts[targetIndex];
                int currentIndex = Shortcuts.IndexOf(shortcut);
                if (currentIndex >= 0 && currentIndex != targetIndex)
                {
                    Shortcuts.Move(currentIndex, targetIndex);
                }
            }

            if (currentSelection != null)
            {
                SelectedShortcut = currentSelection;
            }

            Save();
        }

        private void ConfigureGroupedShortcuts()
        {
            groupedShortcuts = CollectionViewSource.GetDefaultView(model.Shortcuts);
            groupedShortcuts.GroupDescriptions.Clear();
            groupedShortcuts.SortDescriptions.Clear();
            groupedShortcuts.Refresh();
        }

        private void SubscribeToShortcuts(ObservableCollection<Shortcut> shortcuts)
        {
            shortcuts.CollectionChanged -= Shortcuts_CollectionChanged;
            shortcuts.CollectionChanged += Shortcuts_CollectionChanged;

            foreach (Shortcut shortcut in shortcuts)
            {
                shortcut.GroupName = NormalizeGroupName(shortcut.GroupName);
                shortcut.PropertyChanged -= Shortcut_PropertyChanged;
                shortcut.PropertyChanged += Shortcut_PropertyChanged;
                RegisterGroup(shortcut.GroupName);
            }
        }

        private void UnsubscribeFromShortcuts(ObservableCollection<Shortcut> shortcuts)
        {
            shortcuts.CollectionChanged -= Shortcuts_CollectionChanged;
            foreach (Shortcut shortcut in shortcuts)
            {
                shortcut.PropertyChanged -= Shortcut_PropertyChanged;
            }
        }

        private void Shortcuts_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (Shortcut shortcut in e.OldItems)
                {
                    shortcut.PropertyChanged -= Shortcut_PropertyChanged;
                }
            }

            if (e.NewItems != null)
            {
                foreach (Shortcut shortcut in e.NewItems)
                {
                    shortcut.GroupName = NormalizeGroupName(shortcut.GroupName);
                    shortcut.PropertyChanged -= Shortcut_PropertyChanged;
                    shortcut.PropertyChanged += Shortcut_PropertyChanged;
                    RegisterGroup(shortcut.GroupName);
                }
            }

            RefreshGroups();
            Save();
        }

        private void Shortcut_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Shortcut.GroupName) || e.PropertyName == nameof(Shortcut.Name))
            {
                if (sender is Shortcut shortcut)
                {
                    RegisterGroup(shortcut.GroupName);
                }

                RefreshGroups();
                OnPropertyChanged(nameof(SelectedShortcutGroupName));
            }

            if (e.PropertyName == nameof(Shortcut.TypeName) && sender is Shortcut typedShortcut)
            {
                RegisterType(typedShortcut.TypeName);
            }

            Save();
        }

        private void RefreshGroups()
        {
            foreach (Shortcut shortcut in Shortcuts)
            {
                RegisterGroup(shortcut.GroupName);
            }

            groupedShortcuts.Refresh();
            OnPropertyChanged(nameof(GroupedShortcuts));
            OnPropertyChanged(nameof(AvailableGroups));
        }

        private void RegisterGroup(string? groupName)
        {
            model.EnsureGroupState(NormalizeGroupName(groupName));
        }

        private void RegisterType(string? typeName)
        {
            if (!string.IsNullOrWhiteSpace(typeName))
            {
                model.EnsureType(typeName);
            }

            OnPropertyChanged(nameof(AvailableTypes));
        }

        private void AddType(string? typeName)
        {
            string normalized = NormalizeTypeName(typeName);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            RegisterType(normalized);
            Save();
        }

        private void AddTypeFromPrompt(bool assignToSelectedShortcut)
        {
            string proposedName = SelectedShortcut != null && !string.IsNullOrWhiteSpace(SelectedShortcut.TypeName)
                ? SelectedShortcut.TypeName
                : string.Empty;

            string enteredName = Interaction.InputBox("Enter a new type for this fence.", "Add Type", proposedName);
            if (string.IsNullOrWhiteSpace(enteredName))
            {
                return;
            }

            string normalized = NormalizeTypeName(enteredName);
            AddType(normalized);

            if (assignToSelectedShortcut && SelectedShortcut != null)
            {
                MoveSelectedShortcutToType(normalized);
            }
        }

        private static string NormalizeTypeName(string? typeName)
        {
            return string.IsNullOrWhiteSpace(typeName) ? string.Empty : typeName.Trim();
        }

        private static string GetTypeSortKey(string? typeName)
        {
            string normalized = NormalizeTypeName(typeName);
            return string.IsNullOrWhiteSpace(normalized) ? "~~~~" : normalized;
        }

        private static string NormalizeGroupName(string? groupName)
        {
            return string.IsNullOrWhiteSpace(groupName) ? PalisadeModel.DefaultGroupName : groupName.Trim();
        }

        private static string NormalizeDesktopId(string? desktopId)
        {
            return Guid.TryParse(desktopId, out Guid parsedDesktopId) ? parsedDesktopId.ToString() : string.Empty;
        }

        private static List<Shortcut> GetDraggedShortcuts(object? data)
        {
            if (data is Shortcut shortcut)
            {
                return new List<Shortcut> { shortcut };
            }

            if (data is IEnumerable enumerable)
            {
                return enumerable.OfType<Shortcut>().ToList();
            }

            return new List<Shortcut>();
        }

        private static string[] CreateShellDropFiles(IEnumerable<Shortcut> shortcuts)
        {
            string exportDirectory = Path.Combine(Path.GetTempPath(), "Palisades", "DraggedShortcuts", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(exportDirectory);

            List<string> exportedFiles = new();
            HashSet<string> usedNames = new(StringComparer.OrdinalIgnoreCase);

            foreach (Shortcut shortcut in shortcuts)
            {
                string safeBaseName = SanitizeFileName(string.IsNullOrWhiteSpace(shortcut.Name) ? "Shortcut" : shortcut.Name.Trim());
                if (!usedNames.Add(safeBaseName))
                {
                    int suffix = 2;
                    string candidateName;
                    do
                    {
                        candidateName = $"{safeBaseName} ({suffix++})";
                    }
                    while (!usedNames.Add(candidateName));

                    safeBaseName = candidateName;
                }

                if (TryReuseOriginalShortcutFile(shortcut, exportDirectory, safeBaseName, out string reusedShortcutPath))
                {
                    exportedFiles.Add(reusedShortcutPath);
                    continue;
                }

                string? shellIconLocation = ResolveShellIconLocation(shortcut);

                if (shortcut is UrlShortcut || IsWebUrl(shortcut.UriOrFileAction))
                {
                    string urlPath = Path.Combine(exportDirectory, safeBaseName + ".url");
                    List<string> lines = new()
                    {
                        "[InternetShortcut]",
                        $"URL={shortcut.UriOrFileAction}"
                    };

                    if (TrySplitIconLocation(shellIconLocation, out string iconFile, out int iconIndex))
                    {
                        lines.Add($"IconFile={iconFile}");
                        lines.Add($"IconIndex={iconIndex}");
                    }

                    File.WriteAllLines(urlPath, lines);
                    exportedFiles.Add(urlPath);
                    continue;
                }

                string linkPath = Path.Combine(exportDirectory, safeBaseName + ".lnk");
                CreateWindowsShortcutFile(linkPath, shortcut.UriOrFileAction, shellIconLocation);
                exportedFiles.Add(linkPath);
            }

            return exportedFiles.ToArray();
        }

        private static string SanitizeFileName(string fileName)
        {
            char[] invalidCharacters = Path.GetInvalidFileNameChars();
            return string.Concat(fileName.Select(character => invalidCharacters.Contains(character) ? '_' : character));
        }

        private static bool IsWebUrl(string? value)
        {
            return Uri.TryCreate(value, UriKind.Absolute, out Uri? uri)
                && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
        }

        private static bool TryReuseOriginalShortcutFile(Shortcut shortcut, string exportDirectory, string safeBaseName, out string exportedPath)
        {
            exportedPath = string.Empty;

            string? sourceShortcutPath = ResolveOriginalShortcutPath(shortcut);
            if (string.IsNullOrWhiteSpace(sourceShortcutPath) || !File.Exists(sourceShortcutPath))
            {
                return false;
            }

            string extension = Path.GetExtension(sourceShortcutPath);
            if (!string.Equals(extension, ".lnk", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(extension, ".url", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            exportedPath = Path.Combine(exportDirectory, safeBaseName + extension.ToLowerInvariant());
            File.Copy(sourceShortcutPath, exportedPath, true);
            return true;
        }

        private static string? ResolveOriginalShortcutPath(Shortcut shortcut)
        {
            if (!string.IsNullOrWhiteSpace(shortcut.SourceShortcutPath) && File.Exists(shortcut.SourceShortcutPath))
            {
                return shortcut.SourceShortcutPath;
            }

            foreach (string searchRoot in GetShortcutSearchRoots())
            {
                foreach (string extension in new[] { ".lnk", ".url" })
                {
                    try
                    {
                        IEnumerable<string> candidates = Directory.EnumerateFiles(searchRoot, SanitizeFileName(shortcut.Name) + extension, SearchOption.AllDirectories);
                        foreach (string candidate in candidates)
                        {
                            if (ShortcutMatchesTarget(candidate, shortcut))
                            {
                                shortcut.SourceShortcutPath = candidate;
                                return candidate;
                            }
                        }
                    }
                    catch
                    {
                        // Ignore inaccessible folders while searching common shortcut locations.
                    }
                }
            }

            return null;
        }

        private static IEnumerable<string> GetShortcutSearchRoots()
        {
            return new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory),
                Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)
            }
            .Where(path => !string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static bool ShortcutMatchesTarget(string shortcutPath, Shortcut shortcut)
        {
            if (string.Equals(Path.GetExtension(shortcutPath), ".url", StringComparison.OrdinalIgnoreCase))
            {
                string? url = File.ReadLines(shortcutPath).FirstOrDefault(value => value.StartsWith("URL="));
                string candidateUrl = url?.Replace("URL=", string.Empty).Replace("\"", string.Empty).Trim() ?? string.Empty;
                return string.Equals(candidateUrl, shortcut.UriOrFileAction, StringComparison.OrdinalIgnoreCase);
            }

            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                return false;
            }

            object? shell = null;
            object? link = null;
            try
            {
                shell = Activator.CreateInstance(shellType);
                if (shell == null)
                {
                    return false;
                }

                link = shellType.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { shortcutPath });
                string? candidateTarget = link?.GetType().InvokeMember("TargetPath", System.Reflection.BindingFlags.GetProperty, null, link, null) as string;
                return string.Equals(candidateTarget, shortcut.UriOrFileAction, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (link != null && Marshal.IsComObject(link))
                {
                    Marshal.FinalReleaseComObject(link);
                }

                if (shell != null && Marshal.IsComObject(shell))
                {
                    Marshal.FinalReleaseComObject(shell);
                }
            }
        }

        private static string? ResolveShellIconLocation(Shortcut shortcut)
        {
            if (TrySplitIconLocation(shortcut.ShellIconLocation, out string originalIconFile, out int originalIconIndex)
                && (File.Exists(originalIconFile) || Directory.Exists(originalIconFile)))
            {
                return $"{originalIconFile},{originalIconIndex}";
            }

            if (!string.IsNullOrWhiteSpace(shortcut.UriOrFileAction) && (File.Exists(shortcut.UriOrFileAction) || Directory.Exists(shortcut.UriOrFileAction)))
            {
                return $"{shortcut.UriOrFileAction},0";
            }

            string? convertedIconPath = EnsureIcoIconPath(shortcut.IconPath);
            if (!string.IsNullOrWhiteSpace(convertedIconPath) && File.Exists(convertedIconPath))
            {
                return $"{convertedIconPath},0";
            }

            return null;
        }

        private static bool TrySplitIconLocation(string? iconLocation, out string iconFile, out int iconIndex)
        {
            iconFile = string.Empty;
            iconIndex = 0;

            if (string.IsNullOrWhiteSpace(iconLocation))
            {
                return false;
            }

            string trimmed = iconLocation.Trim();
            int separatorIndex = trimmed.LastIndexOf(',');
            if (separatorIndex > 1 && int.TryParse(trimmed[(separatorIndex + 1)..], out int parsedIndex))
            {
                iconFile = trimmed[..separatorIndex].Trim().Trim('"');
                iconIndex = parsedIndex;
            }
            else
            {
                iconFile = trimmed.Trim('"');
            }

            return !string.IsNullOrWhiteSpace(iconFile);
        }

        private static string? EnsureIcoIconPath(string? iconPath)
        {
            if (string.IsNullOrWhiteSpace(iconPath) || !File.Exists(iconPath))
            {
                return null;
            }

            if (string.Equals(Path.GetExtension(iconPath), ".ico", StringComparison.OrdinalIgnoreCase))
            {
                return iconPath;
            }

            string exportIconDirectory = Path.Combine(PDirectory.GetAppDirectory(), "export-icons");
            PDirectory.EnsureExists(exportIconDirectory);

            string safeBaseName = SanitizeFileName(Path.GetFileNameWithoutExtension(iconPath));
            string hash = Math.Abs(StringComparer.OrdinalIgnoreCase.GetHashCode(iconPath)).ToString("X8");
            string icoPath = Path.Combine(exportIconDirectory, $"{safeBaseName}-{hash}.ico");

            if (File.Exists(icoPath) && File.GetLastWriteTimeUtc(icoPath) >= File.GetLastWriteTimeUtc(iconPath))
            {
                return icoPath;
            }

            using System.Drawing.Bitmap bitmap = new(iconPath);
            IntPtr iconHandle = bitmap.GetHicon();
            try
            {
                using System.Drawing.Icon icon = System.Drawing.Icon.FromHandle(iconHandle);
                using FileStream stream = new(icoPath, FileMode.Create, FileAccess.Write, FileShare.None);
                icon.Save(stream);
            }
            finally
            {
                Palisades.Helpers.Native.Bindings.DestroyIcon(iconHandle);
            }

            return icoPath;
        }

        private static void CreateWindowsShortcutFile(string shortcutPath, string targetPath, string? iconLocation)
        {
            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                File.WriteAllText(shortcutPath + ".txt", targetPath ?? string.Empty);
                return;
            }

            object? shell = null;
            object? link = null;

            try
            {
                shell = Activator.CreateInstance(shellType);
                if (shell == null)
                {
                    return;
                }

                link = shellType.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { shortcutPath });
                if (link == null)
                {
                    return;
                }

                Type linkType = link.GetType();
                linkType.InvokeMember("TargetPath", System.Reflection.BindingFlags.SetProperty, null, link, new object[] { targetPath ?? string.Empty });

                string? workingDirectory = !string.IsNullOrWhiteSpace(targetPath) && File.Exists(targetPath)
                    ? Path.GetDirectoryName(targetPath)
                    : null;
                if (!string.IsNullOrWhiteSpace(workingDirectory))
                {
                    linkType.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.SetProperty, null, link, new object[] { workingDirectory });
                }

                if (TrySplitIconLocation(iconLocation, out string iconFile, out int iconIndex)
                    && (File.Exists(iconFile) || Directory.Exists(iconFile)))
                {
                    linkType.InvokeMember("IconLocation", System.Reflection.BindingFlags.SetProperty, null, link, new object[] { $"{iconFile},{iconIndex}" });
                }

                linkType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, link, null);
            }
            finally
            {
                if (link != null && Marshal.IsComObject(link))
                {
                    Marshal.FinalReleaseComObject(link);
                }

                if (shell != null && Marshal.IsComObject(shell))
                {
                    Marshal.FinalReleaseComObject(shell);
                }
            }
        }

        private static string? ResolveDropGroupName(IDropInfo dropInfo)
        {
            if (dropInfo.TargetGroup?.Name != null)
            {
                return dropInfo.TargetGroup.Name.ToString();
            }

            if (dropInfo.TargetItem is Shortcut targetShortcut)
            {
                return targetShortcut.GroupName;
            }

            if (dropInfo.TargetItem is CollectionViewGroup collectionViewGroup)
            {
                return collectionViewGroup.Name?.ToString();
            }

            return null;
        }

        private static string[] GetFileDropPaths(object? data)
        {
            if (data is string[] directPaths)
            {
                return directPaths.Where(path => !string.IsNullOrWhiteSpace(path)).ToArray();
            }

            if (data is IEnumerable<string> enumerablePaths && data is not string)
            {
                return enumerablePaths.Where(path => !string.IsNullOrWhiteSpace(path)).ToArray();
            }

            if (data is IDataObject dataObject && dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                return (dataObject.GetData(DataFormats.FileDrop) as string[] ?? Array.Empty<string>())
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .ToArray();
            }

            return Array.Empty<string>();
        }

        private Shortcut? BuildShortcutFromDroppedPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || (!File.Exists(path) && !Directory.Exists(path)))
            {
                return null;
            }

            string extension = Path.GetExtension(path) ?? string.Empty;
            if (string.Equals(extension, ".lnk", StringComparison.OrdinalIgnoreCase))
            {
                return LnkShortcut.BuildFrom(path, Identifier);
            }

            if (string.Equals(extension, ".url", StringComparison.OrdinalIgnoreCase))
            {
                return UrlShortcut.BuildFrom(path, Identifier);
            }

            string displayName = Directory.Exists(path) ? new DirectoryInfo(path).Name : Shortcut.GetName(path);
            string iconPath = Shortcut.GetIcon(path, Identifier);

            return new LnkShortcut(displayName, iconPath, path)
            {
                ShellIconLocation = path
            };
        }

        private bool ImportDroppedPaths(IEnumerable<string> paths)
        {
            List<Shortcut> importedShortcuts = new();

            foreach (Shortcut? shortcutItem in paths.Select(BuildShortcutFromDroppedPath))
            {
                if (shortcutItem == null)
                {
                    continue;
                }

                Shortcuts.Add(shortcutItem);
                importedShortcuts.Add(shortcutItem);
            }

            if (importedShortcuts.Count == 0)
            {
                return false;
            }

            SelectedShortcut = importedShortcuts.Last();
            RefreshGroups();
            Save();
            return true;
        }

        public bool CanStartDrag(IDragInfo dragInfo)
        {
            return GetDraggedShortcuts(dragInfo.SourceItems ?? dragInfo.Data).Count > 0;
        }

        public void StartDrag(IDragInfo dragInfo)
        {
            defaultDragHandler.StartDrag(dragInfo);

            List<Shortcut> draggedShortcuts = GetDraggedShortcuts(dragInfo.SourceItems ?? dragInfo.Data);
            if (draggedShortcuts.Count == 0)
            {
                return;
            }

            dragInfo.Data = draggedShortcuts.Count == 1 ? draggedShortcuts[0] : draggedShortcuts;

            DataObject dataObject = dragInfo.DataObject as DataObject ?? new DataObject();
            dataObject.SetData(ShortcutDragDataFormat, draggedShortcuts);
            dataObject.SetData(typeof(List<Shortcut>), draggedShortcuts);
            if (draggedShortcuts.Count == 1)
            {
                dataObject.SetData(typeof(Shortcut), draggedShortcuts[0]);
            }

            string[] exportedFiles = CreateShellDropFiles(draggedShortcuts);
            if (exportedFiles.Length > 0)
            {
                dataObject.SetData(DataFormats.FileDrop, exportedFiles);
            }

            dragInfo.DataObject = dataObject;
            dragInfo.Effects = DragDropEffects.Move | DragDropEffects.Copy;
        }

        public void Dropped(IDropInfo dropInfo)
        {
            defaultDragHandler.Dropped(dropInfo);
        }

        public void DragCancelled()
        {
            defaultDragHandler.DragCancelled();
        }

        public bool TryCatchOccurredException(Exception exception)
        {
            return defaultDragHandler.TryCatchOccurredException(exception);
        }

        public void DragDropOperationFinished(DragDropEffects operationResult, IDragInfo dragInfo)
        {
            defaultDragHandler.DragDropOperationFinished(operationResult, dragInfo);
        }

        public void DragOver(IDropInfo dropInfo)
        {
            List<Shortcut> draggedShortcuts = GetDraggedShortcuts(dropInfo.Data);
            if (draggedShortcuts.Count == 0)
            {
                string[] externalPaths = GetFileDropPaths(dropInfo.Data);
                if (externalPaths.Length == 0)
                {
                    dropInfo.Effects = DragDropEffects.None;
                    return;
                }

                dropInfo.Effects = DragDropEffects.Copy;
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.NotHandled = false;
                return;
            }

            if (dropInfo.TargetCollection == null)
            {
                dropInfo.Effects = DragDropEffects.Move;
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.NotHandled = false;
                return;
            }

            defaultDropHandler.DragOver(dropInfo);
            dropInfo.Effects = DragDropEffects.Move;
            dropInfo.DropTargetAdorner = ResolveDropGroupName(dropInfo) == null ? DropTargetAdorners.Insert : DropTargetAdorners.Highlight;
            dropInfo.NotHandled = false;
        }

        public void Drop(IDropInfo dropInfo)
        {
            List<Shortcut> draggedShortcuts = GetDraggedShortcuts(dropInfo.Data);
            if (draggedShortcuts.Count == 0)
            {
                dropInfo.NotHandled = true;
                return;
            }

            string? destinationGroupName = ResolveDropGroupName(dropInfo);

            if (dropInfo.TargetCollection == null)
            {
                foreach (Shortcut shortcut in draggedShortcuts)
                {
                    foreach (PalisadeViewModel sourceViewModel in PalisadesManager.palisades.Values
                                 .Select(window => window.DataContext as PalisadeViewModel)
                                 .OfType<PalisadeViewModel>())
                    {
                        if (ReferenceEquals(sourceViewModel, this))
                        {
                            continue;
                        }

                        if (!sourceViewModel.Shortcuts.Contains(shortcut))
                        {
                            continue;
                        }

                        sourceViewModel.Shortcuts.Remove(shortcut);
                        sourceViewModel.Save();
                        break;
                    }

                    if (!Shortcuts.Contains(shortcut))
                    {
                        Shortcuts.Add(shortcut);
                    }
                }
            }
            else
            {
                defaultDropHandler.Drop(dropInfo);
            }

            if (!string.IsNullOrWhiteSpace(destinationGroupName))
            {
                string normalized = NormalizeGroupName(destinationGroupName);
                foreach (Shortcut shortcut in draggedShortcuts)
                {
                    shortcut.GroupName = normalized;
                }
                RegisterGroup(normalized);
            }

            if (draggedShortcuts.Count == 1)
            {
                SelectedShortcut = draggedShortcuts[0];
            }

            RefreshGroups();
            Save();
        }

        #endregion

        #region Commands
        public ICommand NewPalisadeCommand { get; private set; } = new RelayCommand(() =>
        {
            PalisadesManager.CreatePalisade();
        });

        public ICommand DeletePalisadeCommand { get; private set; } = new RelayCommand<string>((identifier) => PalisadesManager.DeletePalisade(identifier));

        public ICommand HidePalisadeCommand { get; private set; } = new RelayCommand<string>((identifier) => PalisadesManager.HidePalisade(identifier));

        public ICommand JoinPalisadeAsTabCommand
        {
            get
            {
                return new RelayCommand<string>((identifier) => PalisadesManager.JoinPalisadesAsTabs(Identifier, identifier));
            }
        }

        public ICommand SelectJoinedTabCommand
        {
            get
            {
                return new RelayCommand<string>(PalisadesManager.ActivateTabbedFence);
            }
        }

        public ICommand SplitTabbedFenceCommand
        {
            get
            {
                return new RelayCommand(() => PalisadesManager.SplitPalisadeFromTabs(Identifier), () => HasTabbedGroup);
            }
        }

        public ICommand MovePalisadeToDesktopCommand
        {
            get
            {
                return new RelayCommand<string>((desktopId) => PalisadesManager.MovePalisadeToDesktop(Identifier, desktopId), () => VirtualDesktopHelper.HasMultipleDesktops());
            }
        }

        public ICommand ToggleDesktopVisibilityCommand
        {
            get
            {
                return new RelayCommand<string>(ToggleDesktopVisibility, () => VirtualDesktopHelper.HasMultipleDesktops());
            }
        }

        public ICommand MovePalisadeToPreviousDesktopCommand { get; private set; } = new RelayCommand<string>((identifier) => PalisadesManager.MovePalisadeToPreviousDesktop(identifier), () => VirtualDesktopHelper.HasMultipleDesktops());

        public ICommand MovePalisadeToNextDesktopCommand { get; private set; } = new RelayCommand<string>((identifier) => PalisadesManager.MovePalisadeToNextDesktop(identifier), () => VirtualDesktopHelper.HasMultipleDesktops());

        public ICommand ToggleCollapsedStateCommand
        {
            get
            {
                return new RelayCommand(ToggleCollapsedState);
            }
        }

        public ICommand SortByNameAscendingCommand
        {
            get
            {
                return new RelayCommand(() => SortShortcutsByName(false));
            }
        }

        public ICommand SortByNameDescendingCommand
        {
            get
            {
                return new RelayCommand(() => SortShortcutsByName(true));
            }
        }

        public ICommand SortByTypeCommand
        {
            get
            {
                return new RelayCommand(SortShortcutsByType);
            }
        }

        public ICommand AddTypeCommand
        {
            get
            {
                return new RelayCommand(() => AddTypeFromPrompt(false));
            }
        }

        public ICommand AssignSelectedShortcutTypeCommand
        {
            get
            {
                return new RelayCommand<string>((typeName) => MoveSelectedShortcutToType(typeName));
            }
        }

        public ICommand AddTypeAndAssignToSelectedShortcutCommand
        {
            get
            {
                return new RelayCommand(() => AddTypeFromPrompt(true));
            }
        }

        public ICommand ClearSelectedShortcutTypeCommand
        {
            get
            {
                return new RelayCommand(() => MoveSelectedShortcutToType(null));
            }
        }

        public ICommand EditPalisadeCommand { get; private set; } = new RelayCommand<PalisadeViewModel>((viewModel) =>
        {
            EditPalisade edit = new()
            {
                DataContext = viewModel,
                Owner = PalisadesManager.GetPalisade(viewModel.Identifier),
                ShowInTaskbar = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };
            edit.Show();
            edit.Activate();
        });

        public ICommand OpenAboutCommand { get; private set; } = new RelayCommand<PalisadeViewModel>((viewModel) =>
        {
            About about = new()
            {
                DataContext = new AboutViewModel(),
                Owner = PalisadesManager.GetPalisade(viewModel.Identifier)
            };
            about.ShowDialog();
        });

        public ICommand DropShortcut
        {
            get
            {
                return new RelayCommand<DragEventArgs>(DropShortcutsHandler);
            }
        }

        public void DropShortcutsHandler(DragEventArgs dragEventArgs)
        {
            dragEventArgs.Handled = true;
            if (dragEventArgs.Data.GetDataPresent(ShortcutDragDataFormat))
            {
                dragEventArgs.Handled = false;
                return;
            }

            string[] droppedPaths = GetFileDropPaths(dragEventArgs.Data);
            if (droppedPaths.Length == 0)
            {
                dragEventArgs.Handled = false;
                return;
            }

            bool imported = ImportDroppedPaths(droppedPaths);
            dragEventArgs.Effects = imported ? DragDropEffects.Copy : DragDropEffects.None;
        }

        public ICommand ClickShortcut
        {
            get
            {
                return new RelayCommand<Shortcut>(ToggleShortcutSelection);
            }
        }

        public ICommand SelectShortcutForContextMenuCommand
        {
            get
            {
                return new RelayCommand<Shortcut>(SelectShortcutForContextMenu);
            }
        }

        public void SelectShortcutForContextMenu(Shortcut shortcut)
        {
            if (shortcut == null)
            {
                return;
            }

            SelectedShortcut = shortcut;
        }

        public void ToggleShortcutSelection(Shortcut shortcut)
        {
            if (SelectedShortcut == shortcut)
            {
                SelectedShortcut = null;
                return;
            }

            SelectedShortcut = shortcut;
        }

        public ICommand MoveSelectedShortcutToGroupCommand
        {
            get
            {
                return new RelayCommand<string>((groupName) => MoveSelectedShortcutToGroup(groupName), () => HasSelectedShortcut);
            }
        }

        public ICommand RemoveSelectedShortcutCommand
        {
            get
            {
                return new RelayCommand<Shortcut>((shortcut) => RemoveShortcut(shortcut));
            }
        }

        public ICommand UndoRemoveShortcutCommand
        {
            get
            {
                return new RelayCommand(UndoLastShortcutRemoval);
            }
        }

        public ICommand DelKeyPressed
        {
            get
            {
                return new RelayCommand<KeyEventArgs>(HandleShortcutKeyPressed);
            }
        }

        public void HandleShortcutKeyPressed(KeyEventArgs? keyEventArgs)
        {
            if (keyEventArgs == null)
            {
                return;
            }

            Key pressedKey = keyEventArgs.Key == Key.System ? keyEventArgs.SystemKey : keyEventArgs.Key;
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && pressedKey == Key.Z)
            {
                UndoLastShortcutRemoval();
                keyEventArgs.Handled = true;
                return;
            }

            if (pressedKey != Key.Delete)
            {
                return;
            }

            DeleteShortcut();
            keyEventArgs.Handled = true;
        }

        public void DeleteShortcut()
        {
            RemoveShortcut(SelectedShortcut);
        }

        public void RemoveShortcut(Shortcut? shortcut)
        {
            Shortcut? shortcutToRemove = shortcut ?? SelectedShortcut;
            if (shortcutToRemove == null)
            {
                return;
            }

            int removedIndex = Shortcuts.IndexOf(shortcutToRemove);
            if (removedIndex < 0)
            {
                return;
            }

            lastRemovedShortcut = shortcutToRemove;
            lastRemovedShortcutIndex = removedIndex;

            Shortcuts.Remove(shortcutToRemove);
            if (ReferenceEquals(SelectedShortcut, shortcutToRemove))
            {
                SelectedShortcut = null;
            }

            OnPropertyChanged(nameof(CanUndoShortcutRemoval));
            CommandManager.InvalidateRequerySuggested();
        }

        public void UndoLastShortcutRemoval()
        {
            if (lastRemovedShortcut == null || lastRemovedShortcutIndex < 0)
            {
                return;
            }

            if (Shortcuts.Contains(lastRemovedShortcut))
            {
                SelectedShortcut = lastRemovedShortcut;
                lastRemovedShortcut = null;
                lastRemovedShortcutIndex = -1;
                OnPropertyChanged(nameof(CanUndoShortcutRemoval));
                CommandManager.InvalidateRequerySuggested();
                return;
            }

            int restoreIndex = Math.Min(lastRemovedShortcutIndex, Shortcuts.Count);
            Shortcuts.Insert(restoreIndex, lastRemovedShortcut);
            SelectedShortcut = lastRemovedShortcut;
            lastRemovedShortcut = null;
            lastRemovedShortcutIndex = -1;
            OnPropertyChanged(nameof(CanUndoShortcutRemoval));
            CommandManager.InvalidateRequerySuggested();
        }

        /// <summary>
        /// Save asynchronously every 1s if needed.
        /// </summary>
        private void SaveAsync()
        {
            while (true)
            {
                if (shouldSave)
                {
                    string saveDirectory = PDirectory.GetPalisadeDirectory(Identifier);
                    PDirectory.EnsureExists(saveDirectory);
                    using StreamWriter writer = new(Path.Combine(saveDirectory, "state.xml"));
                    XmlSerializer serializer = new(typeof(PalisadeModel), new Type[] { typeof(Shortcut), typeof(LnkShortcut), typeof(UrlShortcut) });
                    serializer.Serialize(writer, this.model);
                    shouldSave = false;
                }
                Thread.Sleep(1000);
            }
        }
        #endregion

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
