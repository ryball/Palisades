
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;
using System.Xml.Serialization;
namespace Palisades.Model
{
    public class PalisadeGroupState
    {
        private string groupName = PalisadeModel.DefaultGroupName;
        private bool isExpanded = true;

        public string GroupName
        {
            get { return string.IsNullOrWhiteSpace(groupName) ? PalisadeModel.DefaultGroupName : groupName; }
            set { groupName = string.IsNullOrWhiteSpace(value) ? PalisadeModel.DefaultGroupName : value.Trim(); }
        }

        public bool IsExpanded { get { return isExpanded; } set { isExpanded = value; } }
    }

    [XmlRoot(Namespace = "io.stouder", ElementName = "PalisadeModel")]
    public class PalisadeModel
    {
        public const string DefaultGroupName = "Ungrouped";
        public const int DefaultHeaderHeight = 60;
        public const int MinHeaderHeight = 36;
        public const int MaxHeaderHeight = 120;
        public const int DefaultIconSize = 48;
        public const int MinIconSize = 32;
        public const int MaxIconSize = 96;
        public const int DefaultFenceHeight = 450;

        private string identifier;
        private string name;
        private string tabGroupId;
        private string activeTabIdentifier;
        private int tabOrder;
        private int fenceX;
        private int fenceY;
        private int width;
        private int height;
        private int headerHeight;
        private int iconSize;
        private bool isCollapsed;
        private bool isLayoutLocked;
        private bool isSearchEnabled;
        private bool showInAltTab;
        private string themePreset;
        private string titleFontFamilyName;
        private string backgroundImagePath;
        private double backgroundImageOpacity;
        private string frameOverlayPath;
        private double frameOverlayOpacity;
        private Color accentColor;
        private ObservableCollection<Shortcut> shortcuts;
        private ObservableCollection<PalisadeGroupState> groupStates;
        private ObservableCollection<string> types;
        private ObservableCollection<string> visibleDesktopIds;
        private Color headerColor;
        private Color bodyColor;
        private Color titleColor;
        private Color labelsColor;

        public PalisadeModel()
        {
            identifier = Guid.NewGuid().ToString();
            name = "No name";
            tabGroupId = identifier;
            activeTabIdentifier = identifier;
            tabOrder = 0;
            headerColor = Color.FromArgb(200, 0, 0, 0);
            bodyColor = Color.FromArgb(120, 0, 0, 0);
            titleColor = Color.FromArgb(255, 255, 255, 255);
            labelsColor = Color.FromArgb(255, 255, 255, 255);
            accentColor = Color.FromArgb(255, 96, 229, 255);
            width = 800;
            height = DefaultFenceHeight;
            headerHeight = DefaultHeaderHeight;
            iconSize = DefaultIconSize;
            isCollapsed = false;
            isLayoutLocked = false;
            isSearchEnabled = true;
            showInAltTab = false;
            themePreset = "Midnight";
            titleFontFamilyName = "Segoe UI";
            backgroundImagePath = string.Empty;
            backgroundImageOpacity = 0.22d;
            frameOverlayPath = string.Empty;
            frameOverlayOpacity = 0.9d;
            shortcuts = new();
            groupStates = new();
            types = new();
            visibleDesktopIds = new();
        }

        public string Identifier { get { return identifier; } set { identifier = value; } }
        public string Name { get { return name; } set { name = value; } }
        public string TabGroupId { get { return string.IsNullOrWhiteSpace(tabGroupId) ? Identifier : tabGroupId; } set { tabGroupId = string.IsNullOrWhiteSpace(value) ? Identifier : value.Trim(); } }
        public string ActiveTabIdentifier { get { return string.IsNullOrWhiteSpace(activeTabIdentifier) ? Identifier : activeTabIdentifier; } set { activeTabIdentifier = string.IsNullOrWhiteSpace(value) ? Identifier : value.Trim(); } }
        public int TabOrder { get { return tabOrder < 0 ? 0 : tabOrder; } set { tabOrder = Math.Max(0, value); } }

        public int FenceX { get { return fenceX; } set { fenceX = value; } }
        public int FenceY { get { return fenceY; } set { fenceY = value; } }

        public int Width { get { return width; } set { width = value; } }
        public int Height { get { return height < HeaderHeight ? DefaultFenceHeight : height; } set { height = Math.Max(value, HeaderHeight); } }
        public int HeaderHeight
        {
            get { return headerHeight < MinHeaderHeight ? DefaultHeaderHeight : headerHeight; }
            set { headerHeight = Math.Clamp(value, MinHeaderHeight, MaxHeaderHeight); }
        }
        public int IconSize
        {
            get { return iconSize < MinIconSize ? DefaultIconSize : iconSize; }
            set { iconSize = Math.Clamp(value, MinIconSize, MaxIconSize); }
        }
        public bool IsCollapsed { get { return isCollapsed; } set { isCollapsed = value; } }
        public bool IsLayoutLocked { get { return isLayoutLocked; } set { isLayoutLocked = value; } }
        public bool IsSearchEnabled { get { return isSearchEnabled; } set { isSearchEnabled = value; } }
        public bool ShowInAltTab { get { return showInAltTab; } set { showInAltTab = value; } }
        public string ThemePreset { get { return string.IsNullOrWhiteSpace(themePreset) ? "Midnight" : themePreset; } set { themePreset = string.IsNullOrWhiteSpace(value) ? "Midnight" : value.Trim(); } }
        public string TitleFontFamilyName { get { return string.IsNullOrWhiteSpace(titleFontFamilyName) ? "Segoe UI" : titleFontFamilyName; } set { titleFontFamilyName = string.IsNullOrWhiteSpace(value) ? "Segoe UI" : value.Trim(); } }
        public string BackgroundImagePath { get { return backgroundImagePath ?? string.Empty; } set { backgroundImagePath = value?.Trim() ?? string.Empty; } }
        public double BackgroundImageOpacity { get { return backgroundImageOpacity < 0.05d ? 0.22d : Math.Clamp(backgroundImageOpacity, 0.05d, 1d); } set { backgroundImageOpacity = Math.Clamp(value, 0.05d, 1d); } }
        public string FrameOverlayPath { get { return frameOverlayPath ?? string.Empty; } set { frameOverlayPath = value?.Trim() ?? string.Empty; } }
        public double FrameOverlayOpacity { get { return frameOverlayOpacity < 0.05d ? 0.9d : Math.Clamp(frameOverlayOpacity, 0.05d, 1d); } set { frameOverlayOpacity = Math.Clamp(value, 0.05d, 1d); } }

        public Color HeaderColor { get { return headerColor; } set { headerColor = value; } }
        public Color BodyColor { get { return bodyColor; } set { bodyColor = value; } }
        public Color TitleColor { get { return titleColor; } set { titleColor = value; } }
        public Color LabelsColor { get { return labelsColor; } set { labelsColor = value; } }
        public Color AccentColor
        {
            get { return accentColor.A == 0 && accentColor.R == 0 && accentColor.G == 0 && accentColor.B == 0 ? Color.FromArgb(255, 96, 229, 255) : accentColor; }
            set { accentColor = value; }
        }

        [XmlArrayItem(typeof(LnkShortcut))]
        [XmlArrayItem(typeof(UrlShortcut))]
        public ObservableCollection<Shortcut> Shortcuts { get { return shortcuts; } set { shortcuts = value ?? new(); } }

        public ObservableCollection<PalisadeGroupState> GroupStates { get { return groupStates; } set { groupStates = value ?? new(); } }

        public ObservableCollection<string> Types { get { return types; } set { types = value ?? new(); } }

        public ObservableCollection<string> VisibleDesktopIds { get { return visibleDesktopIds; } set { visibleDesktopIds = value ?? new(); } }

        public void EnsureDefaults()
        {
            headerHeight = HeaderHeight;
            iconSize = IconSize;
            height = Math.Max(height, HeaderHeight);
            tabGroupId = TabGroupId;
            activeTabIdentifier = ActiveTabIdentifier;
            tabOrder = TabOrder;
            shortcuts ??= new();
            groupStates ??= new();
            types ??= new();
            visibleDesktopIds ??= new();
            themePreset = ThemePreset;
            titleFontFamilyName = TitleFontFamilyName;
            backgroundImagePath = BackgroundImagePath;
            backgroundImageOpacity = BackgroundImageOpacity;
            frameOverlayPath = FrameOverlayPath;
            frameOverlayOpacity = FrameOverlayOpacity;
            accentColor = AccentColor;

            foreach (Shortcut shortcut in shortcuts)
            {
                string normalizedGroup = string.IsNullOrWhiteSpace(shortcut.GroupName) ? DefaultGroupName : shortcut.GroupName.Trim();
                if (!string.Equals(normalizedGroup, DefaultGroupName, StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(shortcut.TypeName))
                    {
                        shortcut.TypeName = normalizedGroup;
                    }

                    shortcut.GroupName = DefaultGroupName;
                }
                else
                {
                    shortcut.GroupName = DefaultGroupName;
                }

                if (string.IsNullOrWhiteSpace(shortcut.ShellIconLocation))
                {
                    shortcut.ShellIconLocation = !string.IsNullOrWhiteSpace(shortcut.UriOrFileAction)
                        ? shortcut.UriOrFileAction
                        : shortcut.IconPath;
                }

                EnsureGroupState(shortcut.GroupName);
                EnsureType(shortcut.TypeName);
            }

            for (int i = types.Count - 1; i >= 0; i--)
            {
                string normalized = NormalizeType(types[i]);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    types.RemoveAt(i);
                    continue;
                }

                types[i] = normalized;
            }

            foreach (PalisadeGroupState state in groupStates.Where(state => state != null))
            {
                state.GroupName = state.GroupName;
            }

            for (int i = visibleDesktopIds.Count - 1; i >= 0; i--)
            {
                string normalizedDesktopId = NormalizeDesktopId(visibleDesktopIds[i]);
                if (string.IsNullOrWhiteSpace(normalizedDesktopId))
                {
                    visibleDesktopIds.RemoveAt(i);
                    continue;
                }

                visibleDesktopIds[i] = normalizedDesktopId;
            }
        }

        public PalisadeGroupState EnsureGroupState(string? groupName)
        {
            string normalized = string.IsNullOrWhiteSpace(groupName) ? DefaultGroupName : groupName.Trim();
            PalisadeGroupState? existing = groupStates.FirstOrDefault(state => string.Equals(state.GroupName, normalized, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            PalisadeGroupState created = new()
            {
                GroupName = normalized,
                IsExpanded = true
            };
            groupStates.Add(created);
            return created;
        }

        public string EnsureType(string? typeName)
        {
            string normalized = NormalizeType(typeName);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            string? existing = types.FirstOrDefault(type => string.Equals(type, normalized, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            types.Add(normalized);
            return normalized;
        }

        private static string NormalizeType(string? typeName)
        {
            return string.IsNullOrWhiteSpace(typeName) ? string.Empty : typeName.Trim();
        }

        private static string NormalizeDesktopId(string? desktopId)
        {
            return Guid.TryParse(desktopId, out Guid parsedDesktopId) ? parsedDesktopId.ToString() : string.Empty;
        }
    }
}
