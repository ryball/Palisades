using GongSolutions.Wpf.DragDrop;
using Microsoft.VisualBasic;
using Microsoft.Win32;
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
using Forms = System.Windows.Forms;

namespace Palisades.ViewModel
{
    public class PalisadeViewModel : INotifyPropertyChanged, IDropTarget, IDragSource
    {
        #region Attributs
        private const string ShortcutDragDataFormat = "Palisades.ShortcutDrag";
        private const string CustomThemePresetName = "Custom";
        private const string DefaultThemePresetName = "Midnight";
        private const string DefaultTitleFontFamilyName = "Segoe UI";
        private static readonly string[] ThemePresetOrder = new[] { "Galactic", DefaultThemePresetName, "Aurora", "Forest", "Sunset", "Paper" };
        private static readonly IReadOnlyDictionary<string, (Color Header, Color Body, Color Title, Color Labels, Color Accent)> ThemePresets = new Dictionary<string, (Color Header, Color Body, Color Title, Color Labels, Color Accent)>(StringComparer.OrdinalIgnoreCase)
        {
            ["Galactic"] = (Color.FromArgb(220, 7, 31, 58), Color.FromArgb(150, 3, 11, 28), Color.FromArgb(255, 226, 249, 255), Color.FromArgb(255, 214, 236, 255), Color.FromArgb(255, 96, 229, 255)),
            [DefaultThemePresetName] = (Color.FromArgb(210, 12, 18, 28), Color.FromArgb(135, 7, 11, 18), Colors.White, Colors.White, Color.FromArgb(255, 134, 195, 255)),
            ["Aurora"] = (Color.FromArgb(220, 34, 56, 89), Color.FromArgb(150, 13, 31, 50), Color.FromArgb(255, 228, 248, 255), Color.FromArgb(255, 216, 240, 255), Color.FromArgb(255, 96, 238, 223)),
            ["Forest"] = (Color.FromArgb(220, 28, 68, 52), Color.FromArgb(145, 18, 42, 32), Color.FromArgb(255, 237, 250, 241), Color.FromArgb(255, 221, 242, 228), Color.FromArgb(255, 120, 235, 173)),
            ["Sunset"] = (Color.FromArgb(220, 111, 46, 52), Color.FromArgb(145, 64, 24, 34), Color.FromArgb(255, 255, 242, 236), Color.FromArgb(255, 255, 228, 214), Color.FromArgb(255, 255, 176, 110)),
            ["Paper"] = (Color.FromArgb(235, 245, 245, 245), Color.FromArgb(205, 255, 255, 255), Color.FromArgb(255, 44, 44, 44), Color.FromArgb(255, 58, 58, 58), Color.FromArgb(255, 74, 145, 216))
        };

        private readonly DefaultDropHandler defaultDropHandler = new();
        private readonly DefaultDragHandler defaultDragHandler = new();
        private readonly PalisadeModel model;

        private ICollectionView groupedShortcuts = null!;
        private volatile bool shouldSave;
        private readonly ObservableCollection<Shortcut> selectedShortcuts = new();
        private readonly List<(Shortcut Shortcut, int Index)> lastRemovedShortcuts = new();
        private readonly Stack<ShortcutHistorySnapshot> undoShortcutHistory = new();
        private readonly Stack<ShortcutHistorySnapshot> redoShortcutHistory = new();
        private string searchQuery = string.Empty;
        private Shortcut? selectedShortcut;
        private bool isHiddenByUser;
        private bool isApplyingThemePreset;

        private sealed class ShortcutHistorySnapshot
        {
            public ShortcutHistorySnapshot(string serializedModel, IReadOnlyList<int> selectedIndices, int primaryIndex)
            {
                SerializedModel = serializedModel;
                SelectedIndices = selectedIndices.ToList();
                PrimaryIndex = primaryIndex;
            }

            public string SerializedModel { get; }
            public List<int> SelectedIndices { get; }
            public int PrimaryIndex { get; }
        }
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

        public int IconSize
        {
            get { return model.IconSize; }
            set
            {
                int normalized = Math.Clamp(value, PalisadeModel.MinIconSize, PalisadeModel.MaxIconSize);
                if (model.IconSize == normalized)
                {
                    return;
                }

                model.IconSize = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShortcutTileWidth));
                OnPropertyChanged(nameof(ShortcutTextMaxWidth));
                Save();
            }
        }

        public double ShortcutTileWidth
        {
            get { return Math.Max(100d, IconSize + 28d); }
        }

        public double ShortcutTextMaxWidth
        {
            get { return Math.Max(84d, IconSize + 20d); }
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

        public string SearchQuery
        {
            get { return searchQuery; }
            set
            {
                string normalized = value ?? string.Empty;
                if (string.Equals(searchQuery, normalized, StringComparison.Ordinal))
                {
                    return;
                }

                searchQuery = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasActiveSearch));
                groupedShortcuts?.Refresh();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool HasActiveSearch
        {
            get { return IsSearchEnabled && !string.IsNullOrWhiteSpace(SearchQuery); }
        }

        public bool IsSearchEnabled
        {
            get { return model.IsSearchEnabled; }
            set
            {
                if (model.IsSearchEnabled == value)
                {
                    return;
                }

                model.IsSearchEnabled = value;
                if (!value && !string.IsNullOrWhiteSpace(searchQuery))
                {
                    searchQuery = string.Empty;
                    OnPropertyChanged(nameof(SearchQuery));
                }

                OnPropertyChanged();
                OnPropertyChanged(nameof(SearchVisibility));
                OnPropertyChanged(nameof(HasActiveSearch));
                groupedShortcuts?.Refresh();
                CommandManager.InvalidateRequerySuggested();
                Save();
            }
        }

        public Visibility SearchVisibility
        {
            get { return IsSearchEnabled ? Visibility.Visible : Visibility.Collapsed; }
        }

        public IEnumerable<string> AvailableThemePresets
        {
            get { return new[] { CustomThemePresetName }.Concat(ThemePresetOrder).ToList(); }
        }

        public IEnumerable<string> AvailableTitleFonts
        {
            get { return Fonts.SystemFontFamilies.Select(font => font.Source).OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase).ToList(); }
        }

        public string SelectedThemePreset
        {
            get { return NormalizeThemePreset(model.ThemePreset); }
            set
            {
                string normalized = NormalizeThemePreset(value);
                if (string.Equals(model.ThemePreset, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                model.ThemePreset = normalized;
                OnPropertyChanged();
                if (!string.Equals(normalized, CustomThemePresetName, StringComparison.OrdinalIgnoreCase))
                {
                    ApplyThemePreset(normalized);
                }
                Save();
            }
        }

        public FontFamily TitleFontFamily
        {
            get
            {
                string fontName = string.IsNullOrWhiteSpace(model.TitleFontFamilyName) ? DefaultTitleFontFamilyName : model.TitleFontFamilyName;
                try
                {
                    return new FontFamily(fontName);
                }
                catch
                {
                    return new FontFamily(DefaultTitleFontFamilyName);
                }
            }
        }

        public string SelectedTitleFontName
        {
            get { return string.IsNullOrWhiteSpace(model.TitleFontFamilyName) ? DefaultTitleFontFamilyName : model.TitleFontFamilyName; }
            set
            {
                string normalized = string.IsNullOrWhiteSpace(value) ? DefaultTitleFontFamilyName : value.Trim();
                if (string.Equals(model.TitleFontFamilyName, normalized, StringComparison.Ordinal))
                {
                    return;
                }

                model.TitleFontFamilyName = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TitleFontFamily));
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public string BackgroundImagePath
        {
            get { return model.BackgroundImagePath; }
            set
            {
                string normalized = value?.Trim() ?? string.Empty;
                if (string.Equals(model.BackgroundImagePath, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                model.BackgroundImagePath = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasBackgroundImage));
                OnPropertyChanged(nameof(BackgroundImageLabel));
                OnPropertyChanged(nameof(BackgroundImageVisibility));
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public string BackgroundImageLabel
        {
            get { return HasBackgroundImage ? Path.GetFileName(model.BackgroundImagePath) : "No fence background image selected."; }
        }

        public string FrameOverlayPath
        {
            get { return model.FrameOverlayPath; }
            set
            {
                string normalized = value?.Trim() ?? string.Empty;
                if (string.Equals(model.FrameOverlayPath, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                model.FrameOverlayPath = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasFrameOverlay));
                OnPropertyChanged(nameof(FrameOverlayLabel));
                OnPropertyChanged(nameof(FrameOverlayVisibility));
                CommandManager.InvalidateRequerySuggested();
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public string FrameOverlayLabel
        {
            get { return HasFrameOverlay ? Path.GetFileName(model.FrameOverlayPath) : "No theme frame overlay selected."; }
        }

        public bool HasFrameOverlay
        {
            get { return !string.IsNullOrWhiteSpace(model.FrameOverlayPath) && File.Exists(model.FrameOverlayPath); }
        }

        public Visibility FrameOverlayVisibility
        {
            get { return HasFrameOverlay ? Visibility.Visible : Visibility.Collapsed; }
        }

        public bool HasBackgroundImage
        {
            get { return !string.IsNullOrWhiteSpace(model.BackgroundImagePath) && File.Exists(model.BackgroundImagePath); }
        }

        public Visibility BackgroundImageVisibility
        {
            get { return HasBackgroundImage ? Visibility.Visible : Visibility.Collapsed; }
        }

        public double BackgroundImageOpacity
        {
            get { return model.BackgroundImageOpacity; }
            set
            {
                double normalized = Math.Clamp(value, 0.05d, 1d);
                if (Math.Abs(model.BackgroundImageOpacity - normalized) < 0.001d)
                {
                    return;
                }

                model.BackgroundImageOpacity = normalized;
                OnPropertyChanged();
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public double FrameOverlayOpacity
        {
            get { return model.FrameOverlayOpacity; }
            set
            {
                double normalized = Math.Clamp(value, 0.05d, 1d);
                if (Math.Abs(model.FrameOverlayOpacity - normalized) < 0.001d)
                {
                    return;
                }

                model.FrameOverlayOpacity = normalized;
                OnPropertyChanged();
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public Color AccentColor
        {
            get { return model.AccentColor; }
            set
            {
                if (model.AccentColor == value)
                {
                    return;
                }

                model.AccentColor = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AccentBrush));
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public SolidColorBrush AccentBrush
        {
            get => new(model.AccentColor);
        }

        public bool ShowInAltTab
        {
            get { return model.ShowInAltTab; }
            set
            {
                if (model.ShowInAltTab == value)
                {
                    return;
                }

                model.ShowInAltTab = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AltTabVisibilityDescription));
                Save();
            }
        }

        public string AltTabVisibilityDescription
        {
            get
            {
                return ShowInAltTab
                    ? "This fence will appear when you switch windows with Alt+Tab."
                    : "This fence stays out of Alt+Tab so it behaves more like part of the desktop.";
            }
        }

        public bool IsLayoutLocked
        {
            get { return model.IsLayoutLocked; }
            set
            {
                if (model.IsLayoutLocked == value)
                {
                    return;
                }

                model.IsLayoutLocked = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanEditShortcuts));
                OnPropertyChanged(nameof(LayoutLockDescription));
                OnPropertyChanged(nameof(WindowResizeMode));
                CommandManager.InvalidateRequerySuggested();
                Save();
            }
        }

        public bool CanEditShortcuts
        {
            get { return !IsLayoutLocked; }
        }

        public string LayoutLockDescription
        {
            get
            {
                return IsLayoutLocked
                    ? "This fence layout is locked. Unlock it to move, resize, drag, rename, or remove shortcuts."
                    : "Lock this fence layout to prevent accidental moves, resizing, drag-and-drop changes, renaming, or removals.";
            }
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
            get { return IsCollapsed || IsLayoutLocked ? ResizeMode.NoResize : ResizeMode.CanResize; }
        }

        private int CollapsedHeight
        {
            get { return Math.Max(HeaderHeight, PalisadeModel.MinHeaderHeight); }
        }

        public Color HeaderColor
        {
            get { return model.HeaderColor; }
            set
            {
                if (model.HeaderColor == value)
                {
                    return;
                }

                model.HeaderColor = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HeaderBrush));
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public SolidColorBrush HeaderBrush
        {
            get => new(model.HeaderColor);
        }

        public Color BodyColor
        {
            get { return model.BodyColor; }
            set
            {
                if (model.BodyColor == value)
                {
                    return;
                }

                model.BodyColor = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(BodyBrush));
                MarkThemePresetAsCustom();
                Save();
            }
        }

        public SolidColorBrush BodyBrush
        {
            get => new(model.BodyColor);
        }

        public SolidColorBrush TitleColor
        {
            get => new(model.TitleColor);
            set
            {
                Color color = value?.Color ?? Colors.White;
                if (model.TitleColor == color)
                {
                    return;
                }

                model.TitleColor = color;
                OnPropertyChanged();
                MarkThemePresetAsCustom();
                Save();
            }
        }
        public SolidColorBrush LabelsColor
        {
            get => new(model.LabelsColor);
            set
            {
                Color color = value?.Color ?? Colors.White;
                if (model.LabelsColor == color)
                {
                    return;
                }

                model.LabelsColor = color;
                OnPropertyChanged();
                MarkThemePresetAsCustom();
                Save();
            }
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
            private set
            {
                selectedShortcut = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasSelectedShortcut));
                OnPropertyChanged(nameof(HasMultipleSelectedShortcuts));
                OnPropertyChanged(nameof(SelectedShortcutCount));
                OnPropertyChanged(nameof(RemoveSelectedMenuHeader));
                OnPropertyChanged(nameof(SelectedShortcutGroupName));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool HasSelectedShortcut
        {
            get { return selectedShortcuts.Count > 0; }
        }

        public bool HasMultipleSelectedShortcuts
        {
            get { return selectedShortcuts.Count > 1; }
        }

        public int SelectedShortcutCount
        {
            get { return selectedShortcuts.Count; }
        }

        public string RemoveSelectedMenuHeader
        {
            get { return HasMultipleSelectedShortcuts ? $"Remove selected ({SelectedShortcutCount})" : "Remove selected"; }
        }

        public bool CanUndoShortcutRemoval
        {
            get { return undoShortcutHistory.Count > 0 || lastRemovedShortcuts.Count > 0; }
        }

        public bool CanRedoShortcutHistory
        {
            get { return redoShortcutHistory.Count > 0; }
        }

        public string SelectedShortcutGroupName
        {
            get { return SelectedShortcut?.GroupName ?? string.Empty; }
            set
            {
                IReadOnlyList<Shortcut> targets = GetSelectedShortcutSnapshot();
                if (targets.Count == 0)
                {
                    return;
                }

                string normalized = NormalizeGroupName(value);
                bool changed = false;
                foreach (Shortcut shortcut in targets)
                {
                    if (string.Equals(shortcut.GroupName, normalized, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    shortcut.GroupName = normalized;
                    changed = true;
                }

                if (!changed)
                {
                    return;
                }

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
            if (IsLayoutLocked || !HasSelectedShortcut)
            {
                return;
            }

            string normalized = NormalizeGroupName(groupName);
            if (GetSelectedShortcutSnapshot().All(shortcut => string.Equals(shortcut.GroupName, normalized, StringComparison.Ordinal)))
            {
                return;
            }

            RecordShortcutHistorySnapshot();
            SelectedShortcutGroupName = normalized;
        }

        public void MoveSelectedShortcutToType(string? typeName)
        {
            if (IsLayoutLocked || !HasSelectedShortcut)
            {
                return;
            }

            string normalized = NormalizeTypeName(typeName);
            List<Shortcut> targets = GetSelectedShortcutSnapshot()
                .Where(shortcut => !string.Equals(shortcut.TypeName, normalized, StringComparison.Ordinal))
                .ToList();
            if (targets.Count == 0)
            {
                return;
            }

            RecordShortcutHistorySnapshot();
            foreach (Shortcut shortcut in targets)
            {
                shortcut.TypeName = normalized;
            }

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
            if (orderedShortcuts.Count != Shortcuts.Count || orderedShortcuts.SequenceEqual(Shortcuts))
            {
                return;
            }

            RecordShortcutHistorySnapshot();
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
            groupedShortcuts.Filter = FilterShortcut;
            groupedShortcuts.Refresh();
        }

        private bool FilterShortcut(object candidate)
        {
            if (candidate is not Shortcut shortcut)
            {
                return false;
            }

            if (!IsSearchEnabled || string.IsNullOrWhiteSpace(SearchQuery))
            {
                return true;
            }

            string search = SearchQuery.Trim();
            return ContainsSearchValue(shortcut.Name, search)
                || ContainsSearchValue(shortcut.UriOrFileAction, search)
                || ContainsSearchValue(shortcut.TypeName, search)
                || ContainsSearchValue(shortcut.GroupName, search);
        }

        private static bool ContainsSearchValue(string? value, string search)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.IndexOf(search, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private static string NormalizeThemePreset(string? themePreset)
        {
            if (string.IsNullOrWhiteSpace(themePreset))
            {
                return DefaultThemePresetName;
            }

            string normalized = themePreset.Trim();
            if (string.Equals(normalized, CustomThemePresetName, StringComparison.OrdinalIgnoreCase))
            {
                return CustomThemePresetName;
            }

            return ThemePresets.ContainsKey(normalized) ? normalized : DefaultThemePresetName;
        }

        private void ApplyThemePreset(string themePreset)
        {
            string normalized = NormalizeThemePreset(themePreset);
            if (!ThemePresets.TryGetValue(normalized, out (Color Header, Color Body, Color Title, Color Labels, Color Accent) palette))
            {
                return;
            }

            isApplyingThemePreset = true;
            try
            {
                model.ThemePreset = normalized;
                model.HeaderColor = palette.Header;
                model.BodyColor = palette.Body;
                model.TitleColor = palette.Title;
                model.LabelsColor = palette.Labels;
                model.AccentColor = palette.Accent;
            }
            finally
            {
                isApplyingThemePreset = false;
            }

            OnPropertyChanged(nameof(SelectedThemePreset));
            OnPropertyChanged(nameof(HeaderColor));
            OnPropertyChanged(nameof(BodyColor));
            OnPropertyChanged(nameof(TitleColor));
            OnPropertyChanged(nameof(LabelsColor));
            OnPropertyChanged(nameof(AccentColor));
            OnPropertyChanged(nameof(AccentBrush));
        }

        private void MarkThemePresetAsCustom()
        {
            if (isApplyingThemePreset || string.Equals(model.ThemePreset, CustomThemePresetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            model.ThemePreset = CustomThemePresetName;
            OnPropertyChanged(nameof(SelectedThemePreset));
        }

        private void ChooseThemeColor(string? colorTarget)
        {
            Color currentColor = colorTarget switch
            {
                "Header" => HeaderColor,
                "Body" => BodyColor,
                "Title" => TitleColor.Color,
                "Labels" => LabelsColor.Color,
                "Accent" => AccentColor,
                _ => Colors.White
            };

            Forms.ColorDialog dialog = new()
            {
                AllowFullOpen = true,
                FullOpen = true,
                Color = System.Drawing.Color.FromArgb(currentColor.A, currentColor.R, currentColor.G, currentColor.B)
            };

            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                return;
            }

            Color selected = Color.FromArgb(currentColor.A, dialog.Color.R, dialog.Color.G, dialog.Color.B);
            switch (colorTarget)
            {
                case "Header":
                    HeaderColor = selected;
                    break;
                case "Body":
                    BodyColor = selected;
                    break;
                case "Title":
                    TitleColor = new SolidColorBrush(Color.FromArgb(255, dialog.Color.R, dialog.Color.G, dialog.Color.B));
                    break;
                case "Labels":
                    LabelsColor = new SolidColorBrush(Color.FromArgb(255, dialog.Color.R, dialog.Color.G, dialog.Color.B));
                    break;
                case "Accent":
                    AccentColor = Color.FromArgb(255, dialog.Color.R, dialog.Color.G, dialog.Color.B);
                    break;
            }
        }

        private void SelectBackgroundImage()
        {
            OpenFileDialog dialog = new()
            {
                Title = "Choose fence background image",
                Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|All files|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog() == true)
            {
                BackgroundImagePath = dialog.FileName;
            }
        }

        private void ClearBackgroundImage()
        {
            BackgroundImagePath = string.Empty;
        }

        private void SelectFrameOverlay()
        {
            OpenFileDialog dialog = new()
            {
                Title = "Choose theme frame overlay image",
                Filter = "PNG files|*.png|Image files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|All files|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog() == true)
            {
                FrameOverlayPath = dialog.FileName;
            }
        }

        private void ClearFrameOverlay()
        {
            FrameOverlayPath = string.Empty;
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
                    shortcut.IsSelected = false;
                    shortcut.PropertyChanged -= Shortcut_PropertyChanged;
                    selectedShortcuts.Remove(shortcut);
                }
            }

            if (e.NewItems != null)
            {
                foreach (Shortcut shortcut in e.NewItems)
                {
                    shortcut.GroupName = NormalizeGroupName(shortcut.GroupName);
                    shortcut.IsSelected = false;
                    shortcut.PropertyChanged -= Shortcut_PropertyChanged;
                    shortcut.PropertyChanged += Shortcut_PropertyChanged;
                    RegisterGroup(shortcut.GroupName);
                }
            }

            if (selectedShortcut != null && !selectedShortcuts.Contains(selectedShortcut))
            {
                SelectedShortcut = selectedShortcuts.LastOrDefault();
            }
            else
            {
                OnPropertyChanged(nameof(HasSelectedShortcut));
                OnPropertyChanged(nameof(HasMultipleSelectedShortcuts));
                OnPropertyChanged(nameof(SelectedShortcutCount));
                OnPropertyChanged(nameof(RemoveSelectedMenuHeader));
            }

            RefreshGroups();
            Save();
        }

        private void Shortcut_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Shortcut.PendingName) || e.PropertyName == nameof(Shortcut.IsRenaming) || e.PropertyName == nameof(Shortcut.IsSelected))
            {
                return;
            }

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
            OnPropertyChanged(nameof(AvailableGroups));
        }

        private IReadOnlyList<Shortcut> GetSelectedShortcutSnapshot()
        {
            return selectedShortcuts.Where(shortcut => Shortcuts.Contains(shortcut)).ToList();
        }

        private void ApplyShortcutSelection(IEnumerable<Shortcut> shortcuts, Shortcut? primaryShortcut = null)
        {
            HashSet<Shortcut> selectionSet = shortcuts
                .Where(shortcut => shortcut != null && Shortcuts.Contains(shortcut))
                .ToHashSet();

            selectedShortcuts.Clear();
            foreach (Shortcut shortcut in Shortcuts)
            {
                bool isSelected = selectionSet.Contains(shortcut);
                shortcut.IsSelected = isSelected;
                if (isSelected)
                {
                    selectedShortcuts.Add(shortcut);
                }
            }

            Shortcut? resolvedPrimary = primaryShortcut != null && selectionSet.Contains(primaryShortcut)
                ? primaryShortcut
                : selectedShortcuts.LastOrDefault();
            SelectedShortcut = resolvedPrimary;
        }

        private void RecordShortcutHistorySnapshot()
        {
            undoShortcutHistory.Push(CaptureShortcutHistorySnapshot());
            redoShortcutHistory.Clear();
            lastRemovedShortcuts.Clear();
            OnPropertyChanged(nameof(CanUndoShortcutRemoval));
            OnPropertyChanged(nameof(CanRedoShortcutHistory));
            CommandManager.InvalidateRequerySuggested();
        }

        private ShortcutHistorySnapshot CaptureShortcutHistorySnapshot()
        {
            using StringWriter writer = new();
            XmlSerializer serializer = new(typeof(PalisadeModel), new Type[] { typeof(Shortcut), typeof(LnkShortcut), typeof(UrlShortcut) });
            serializer.Serialize(writer, model);

            List<int> selectedIndices = GetSelectedShortcutSnapshot()
                .Select(shortcut => Shortcuts.IndexOf(shortcut))
                .Where(index => index >= 0)
                .Distinct()
                .OrderBy(index => index)
                .ToList();

            int primaryIndex = SelectedShortcut != null ? Shortcuts.IndexOf(SelectedShortcut) : -1;
            return new ShortcutHistorySnapshot(writer.ToString(), selectedIndices, primaryIndex);
        }

        private void RestoreShortcutHistorySnapshot(ShortcutHistorySnapshot snapshot)
        {
            using StringReader reader = new(snapshot.SerializedModel);
            XmlSerializer serializer = new(typeof(PalisadeModel), new Type[] { typeof(Shortcut), typeof(LnkShortcut), typeof(UrlShortcut) });
            if (serializer.Deserialize(reader) is not PalisadeModel restoredModel)
            {
                return;
            }

            restoredModel.EnsureDefaults();
            model.GroupStates = new ObservableCollection<PalisadeGroupState>(
                restoredModel.GroupStates.Select(state => new PalisadeGroupState
                {
                    GroupName = state.GroupName,
                    IsExpanded = state.IsExpanded
                }));
            model.Types = new ObservableCollection<string>(restoredModel.Types);
            Shortcuts = new ObservableCollection<Shortcut>(restoredModel.Shortcuts ?? new ObservableCollection<Shortcut>());

            List<Shortcut> restoredSelection = snapshot.SelectedIndices
                .Where(index => index >= 0 && index < Shortcuts.Count)
                .Select(index => Shortcuts[index])
                .ToList();
            Shortcut? primaryShortcut = snapshot.PrimaryIndex >= 0 && snapshot.PrimaryIndex < Shortcuts.Count
                ? Shortcuts[snapshot.PrimaryIndex]
                : restoredSelection.LastOrDefault();
            ApplyShortcutSelection(restoredSelection, primaryShortcut);

            RefreshGroups();
            Save();
            OnPropertyChanged(nameof(AvailableTypes));
            OnPropertyChanged(nameof(CanUndoShortcutRemoval));
            OnPropertyChanged(nameof(CanRedoShortcutHistory));
            CommandManager.InvalidateRequerySuggested();
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
            if (IsLayoutLocked)
            {
                return false;
            }

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
            if (IsLayoutLocked)
            {
                dropInfo.Effects = DragDropEffects.None;
                dropInfo.NotHandled = false;
                return;
            }

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
            if (IsLayoutLocked)
            {
                dropInfo.NotHandled = false;
                return;
            }

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
                return new RelayCommand(() => AddTypeFromPrompt(false), () => CanEditShortcuts);
            }
        }

        public ICommand AssignSelectedShortcutTypeCommand
        {
            get
            {
                return new RelayCommand<string>((typeName) => MoveSelectedShortcutToType(typeName), () => CanEditShortcuts && HasSelectedShortcut);
            }
        }

        public ICommand AddTypeAndAssignToSelectedShortcutCommand
        {
            get
            {
                return new RelayCommand(() => AddTypeFromPrompt(true), () => CanEditShortcuts && HasSelectedShortcut);
            }
        }

        public ICommand ClearSelectedShortcutTypeCommand
        {
            get
            {
                return new RelayCommand(() => MoveSelectedShortcutToType(null), () => CanEditShortcuts && HasSelectedShortcut);
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
            if (IsLayoutLocked)
            {
                dragEventArgs.Effects = DragDropEffects.None;
                return;
            }
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

            if (selectedShortcuts.Contains(shortcut))
            {
                SelectedShortcut = shortcut;
                return;
            }

            ApplyShortcutSelection(new[] { shortcut }, shortcut);
        }

        public void ToggleShortcutSelection(Shortcut shortcut)
        {
            if (shortcut == null)
            {
                return;
            }

            bool extendSelection = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            if (!extendSelection)
            {
                ApplyShortcutSelection(new[] { shortcut }, shortcut);
                return;
            }

            List<Shortcut> updatedSelection = GetSelectedShortcutSnapshot().ToList();
            if (updatedSelection.Contains(shortcut))
            {
                updatedSelection.Remove(shortcut);
            }
            else
            {
                updatedSelection.Add(shortcut);
            }

            ApplyShortcutSelection(updatedSelection, updatedSelection.Contains(shortcut) ? shortcut : updatedSelection.LastOrDefault());
        }

        public ICommand BeginRenameShortcutCommand
        {
            get
            {
                return new RelayCommand<Shortcut>(BeginRenameShortcut);
            }
        }

        public void BeginRenameShortcut(Shortcut? shortcut)
        {
            if (shortcut == null || IsLayoutLocked || !Shortcuts.Contains(shortcut))
            {
                return;
            }

            foreach (Shortcut existingShortcut in Shortcuts.Where(item => !ReferenceEquals(item, shortcut) && item.IsRenaming))
            {
                existingShortcut.PendingName = existingShortcut.Name;
                existingShortcut.IsRenaming = false;
            }

            SelectedShortcut = shortcut;
            shortcut.PendingName = shortcut.Name;
            shortcut.IsRenaming = true;
        }

        public void CommitRenameShortcut(Shortcut? shortcut)
        {
            if (shortcut == null)
            {
                return;
            }

            string originalName = shortcut.Name;
            string normalized = string.IsNullOrWhiteSpace(shortcut.PendingName) ? originalName : shortcut.PendingName.Trim();
            shortcut.IsRenaming = false;
            shortcut.PendingName = normalized;

            if (string.Equals(originalName, normalized, StringComparison.CurrentCulture))
            {
                return;
            }

            RecordShortcutHistorySnapshot();
            shortcut.Name = normalized;
            RefreshGroups();
            Save();
        }

        public void CancelRenameShortcut(Shortcut? shortcut)
        {
            if (shortcut == null)
            {
                return;
            }

            shortcut.PendingName = shortcut.Name;
            shortcut.IsRenaming = false;
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

        public ICommand RemoveAllSelectedShortcutsCommand
        {
            get
            {
                return new RelayCommand(RemoveSelectedShortcuts, () => CanEditShortcuts && HasSelectedShortcut);
            }
        }

        public ICommand UndoRemoveShortcutCommand
        {
            get
            {
                return new RelayCommand(UndoLastShortcutRemoval);
            }
        }

        public ICommand ClearSearchCommand
        {
            get
            {
                return new RelayCommand(() => SearchQuery = string.Empty, () => HasActiveSearch);
            }
        }

        public ICommand ChooseThemeColorCommand
        {
            get
            {
                return new RelayCommand<string>(ChooseThemeColor);
            }
        }

        public ICommand SelectBackgroundImageCommand
        {
            get
            {
                return new RelayCommand(SelectBackgroundImage);
            }
        }

        public ICommand ClearBackgroundImageCommand
        {
            get
            {
                return new RelayCommand(ClearBackgroundImage, () => HasBackgroundImage);
            }
        }

        public ICommand SelectFrameOverlayCommand
        {
            get
            {
                return new RelayCommand(SelectFrameOverlay);
            }
        }

        public ICommand ClearFrameOverlayCommand
        {
            get
            {
                return new RelayCommand(ClearFrameOverlay, () => HasFrameOverlay);
            }
        }

        public ICommand RedoShortcutCommand
        {
            get
            {
                return new RelayCommand(RedoLastShortcutChange, () => CanRedoShortcutHistory);
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

            if (Keyboard.FocusedElement is System.Windows.Controls.Primitives.TextBoxBase)
            {
                return;
            }

            Key pressedKey = keyEventArgs.Key == Key.System ? keyEventArgs.SystemKey : keyEventArgs.Key;
            bool isControlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            bool isShiftPressed = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;

            if (isControlPressed && (pressedKey == Key.Y || (isShiftPressed && pressedKey == Key.Z)))
            {
                RedoLastShortcutChange();
                keyEventArgs.Handled = true;
                return;
            }

            if (isControlPressed && pressedKey == Key.Z)
            {
                UndoLastShortcutRemoval();
                keyEventArgs.Handled = true;
                return;
            }

            if (isControlPressed && pressedKey == Key.A)
            {
                ApplyShortcutSelection(Shortcuts, SelectedShortcut ?? Shortcuts.FirstOrDefault());
                keyEventArgs.Handled = true;
                return;
            }

            if (pressedKey == Key.Escape)
            {
                ApplyShortcutSelection(Array.Empty<Shortcut>());
                keyEventArgs.Handled = true;
                return;
            }

            if (pressedKey == Key.F2)
            {
                BeginRenameShortcut(SelectedShortcut);
                keyEventArgs.Handled = true;
                return;
            }

            if (pressedKey == Key.Enter)
            {
                LaunchSelectedShortcut();
                keyEventArgs.Handled = true;
                return;
            }

            if (pressedKey == Key.Left || pressedKey == Key.Right || pressedKey == Key.Up || pressedKey == Key.Down)
            {
                MoveKeyboardSelection(pressedKey, (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control);
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
            RemoveSelectedShortcuts();
        }

        public void RemoveSelectedShortcuts()
        {
            RemoveShortcuts(GetSelectedShortcutSnapshot());
        }

        public void RemoveShortcut(Shortcut? shortcut)
        {
            if (shortcut == null)
            {
                RemoveSelectedShortcuts();
                return;
            }

            RemoveShortcuts(new[] { shortcut });
        }

        private void RemoveShortcuts(IEnumerable<Shortcut> shortcutsToRemove)
        {
            if (IsLayoutLocked)
            {
                return;
            }

            List<(Shortcut Shortcut, int Index)> removalBatch = shortcutsToRemove
                .Where(shortcut => shortcut != null)
                .Distinct()
                .Select(shortcut => (Shortcut: shortcut, Index: Shortcuts.IndexOf(shortcut)))
                .Where(entry => entry.Index >= 0)
                .OrderByDescending(entry => entry.Index)
                .ToList();

            if (removalBatch.Count == 0)
            {
                return;
            }

            RecordShortcutHistorySnapshot();
            lastRemovedShortcuts.Clear();
            lastRemovedShortcuts.AddRange(removalBatch.OrderBy(entry => entry.Index));

            foreach ((Shortcut shortcut, _) in removalBatch)
            {
                shortcut.IsSelected = false;
                Shortcuts.Remove(shortcut);
                selectedShortcuts.Remove(shortcut);
            }

            SelectedShortcut = selectedShortcuts.LastOrDefault();
            OnPropertyChanged(nameof(CanUndoShortcutRemoval));
            OnPropertyChanged(nameof(CanRedoShortcutHistory));
            CommandManager.InvalidateRequerySuggested();
        }

        public void UndoLastShortcutRemoval()
        {
            if (undoShortcutHistory.Count > 0)
            {
                redoShortcutHistory.Push(CaptureShortcutHistorySnapshot());
                ShortcutHistorySnapshot snapshot = undoShortcutHistory.Pop();
                RestoreShortcutHistorySnapshot(snapshot);
                lastRemovedShortcuts.Clear();
                return;
            }

            if (lastRemovedShortcuts.Count == 0)
            {
                return;
            }

            redoShortcutHistory.Push(CaptureShortcutHistorySnapshot());
            foreach ((Shortcut shortcut, int index) in lastRemovedShortcuts.OrderBy(entry => entry.Index))
            {
                if (Shortcuts.Contains(shortcut))
                {
                    continue;
                }

                int restoreIndex = Math.Min(index, Shortcuts.Count);
                Shortcuts.Insert(restoreIndex, shortcut);
            }

            ApplyShortcutSelection(lastRemovedShortcuts.Select(entry => entry.Shortcut), lastRemovedShortcuts.Last().Shortcut);
            lastRemovedShortcuts.Clear();
            Save();
            OnPropertyChanged(nameof(CanUndoShortcutRemoval));
            OnPropertyChanged(nameof(CanRedoShortcutHistory));
            CommandManager.InvalidateRequerySuggested();
        }

        public void RedoLastShortcutChange()
        {
            if (redoShortcutHistory.Count == 0)
            {
                return;
            }

            undoShortcutHistory.Push(CaptureShortcutHistorySnapshot());
            ShortcutHistorySnapshot snapshot = redoShortcutHistory.Pop();
            RestoreShortcutHistorySnapshot(snapshot);
            lastRemovedShortcuts.Clear();
        }

        private void LaunchSelectedShortcut()
        {
            if (SelectedShortcut == null || string.IsNullOrWhiteSpace(SelectedShortcut.UriOrFileAction))
            {
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = SelectedShortcut.UriOrFileAction,
                    UseShellExecute = true
                });
            }
            catch
            {
                // Ignore launch failures so keyboard navigation never crashes the app.
            }
        }

        private void MoveKeyboardSelection(Key directionKey, bool extendSelection)
        {
            if (Shortcuts.Count == 0)
            {
                return;
            }

            int currentIndex = SelectedShortcut != null ? Shortcuts.IndexOf(SelectedShortcut) : -1;
            if (currentIndex < 0)
            {
                ApplyShortcutSelection(new[] { Shortcuts[0] }, Shortcuts[0]);
                return;
            }

            int step = directionKey == Key.Left || directionKey == Key.Up ? -1 : 1;
            int nextIndex = Math.Clamp(currentIndex + step, 0, Shortcuts.Count - 1);
            Shortcut nextShortcut = Shortcuts[nextIndex];

            if (!extendSelection)
            {
                ApplyShortcutSelection(new[] { nextShortcut }, nextShortcut);
                return;
            }

            List<Shortcut> updatedSelection = GetSelectedShortcutSnapshot().ToList();
            if (!updatedSelection.Contains(nextShortcut))
            {
                updatedSelection.Add(nextShortcut);
            }

            ApplyShortcutSelection(updatedSelection, nextShortcut);
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
