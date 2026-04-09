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
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Xml.Serialization;

namespace Palisades.ViewModel
{
    public class PalisadeViewModel : INotifyPropertyChanged, IDropTarget
    {
        #region Attributs
        private readonly DefaultDropHandler defaultDropHandler = new();
        private readonly PalisadeModel model;

        private ICollectionView groupedShortcuts = null!;
        private volatile bool shouldSave;
        private Shortcut? selectedShortcut;
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
            get { return IsCollapsed ? "+" : "−"; }
        }

        public bool SupportsMultipleDesktops
        {
            get { return VirtualDesktopHelper.HasMultipleDesktops(); }
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
                Save();
            }
        }

        public bool HasSelectedShortcut
        {
            get { return SelectedShortcut != null; }
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

        public void DragOver(IDropInfo dropInfo)
        {
            List<Shortcut> draggedShortcuts = GetDraggedShortcuts(dropInfo.Data);
            if (draggedShortcuts.Count == 0)
            {
                dropInfo.Effects = DragDropEffects.None;
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

            defaultDropHandler.Drop(dropInfo);

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
            if (!dragEventArgs.Data.GetDataPresent(DataFormats.FileDrop))
            {
                dragEventArgs.Handled = false;
                return;
            }

            string[] shortcuts = (string[])dragEventArgs.Data.GetData(DataFormats.FileDrop);
            foreach (string shortcut in shortcuts)
            {
                string? extension = Path.GetExtension(shortcut);

                if (extension == null)
                {
                    continue;
                }

                if (extension == ".lnk")
                {
                    Shortcut? shortcutItem = LnkShortcut.BuildFrom(shortcut, Identifier);
                    if (shortcutItem != null)
                    {
                        Shortcuts.Add(shortcutItem);
                    }
                }
                if (extension == ".url")
                {
                    Shortcut? shortcutItem = UrlShortcut.BuildFrom(shortcut, Identifier);
                    if (shortcutItem != null)
                    {
                        Shortcuts.Add(shortcutItem);
                    }
                }
            }
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
            if (pressedKey != Key.Delete)
            {
                return;
            }

            DeleteShortcut();
            keyEventArgs.Handled = true;
        }

        public void DeleteShortcut()
        {
            if (SelectedShortcut == null)
            {
                return;
            }

            Shortcuts.Remove(SelectedShortcut);
            SelectedShortcut = null;
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
