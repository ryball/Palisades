
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
        public const int DefaultFenceHeight = 450;

        private string identifier;
        private string name;
        private int fenceX;
        private int fenceY;
        private int width;
        private int height;
        private int headerHeight;
        private bool isCollapsed;
        private ObservableCollection<Shortcut> shortcuts;
        private ObservableCollection<PalisadeGroupState> groupStates;
        private ObservableCollection<string> types;
        private Color headerColor;
        private Color bodyColor;
        private Color titleColor;
        private Color labelsColor;

        public PalisadeModel()
        {
            identifier = Guid.NewGuid().ToString();
            name = "No name";
            headerColor = Color.FromArgb(200, 0, 0, 0);
            bodyColor = Color.FromArgb(120, 0, 0, 0);
            titleColor = Color.FromArgb(255, 255, 255, 255);
            labelsColor = Color.FromArgb(255, 255, 255, 255);
            width = 800;
            height = DefaultFenceHeight;
            headerHeight = DefaultHeaderHeight;
            isCollapsed = false;
            shortcuts = new();
            groupStates = new();
            types = new();
        }

        public string Identifier { get { return identifier; } set { identifier = value; } }
        public string Name { get { return name; } set { name = value; } }

        public int FenceX { get { return fenceX; } set { fenceX = value; } }
        public int FenceY { get { return fenceY; } set { fenceY = value; } }

        public int Width { get { return width; } set { width = value; } }
        public int Height { get { return height < HeaderHeight ? DefaultFenceHeight : height; } set { height = Math.Max(value, HeaderHeight); } }
        public int HeaderHeight
        {
            get { return headerHeight < MinHeaderHeight ? DefaultHeaderHeight : headerHeight; }
            set { headerHeight = Math.Clamp(value, MinHeaderHeight, MaxHeaderHeight); }
        }
        public bool IsCollapsed { get { return isCollapsed; } set { isCollapsed = value; } }

        public Color HeaderColor { get { return headerColor; } set { headerColor = value; } }
        public Color BodyColor { get { return bodyColor; } set { bodyColor = value; } }
        public Color TitleColor { get { return titleColor; } set { titleColor = value; } }
        public Color LabelsColor { get { return labelsColor; } set { labelsColor = value; } }

        [XmlArrayItem(typeof(LnkShortcut))]
        [XmlArrayItem(typeof(UrlShortcut))]
        public ObservableCollection<Shortcut> Shortcuts { get { return shortcuts; } set { shortcuts = value ?? new(); } }

        public ObservableCollection<PalisadeGroupState> GroupStates { get { return groupStates; } set { groupStates = value ?? new(); } }

        public ObservableCollection<string> Types { get { return types; } set { types = value ?? new(); } }

        public void EnsureDefaults()
        {
            headerHeight = HeaderHeight;
            height = Math.Max(height, HeaderHeight);
            shortcuts ??= new();
            groupStates ??= new();
            types ??= new();

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
    }
}
