using Palisades.Helpers;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.CompilerServices;
using System.Xml.Serialization;

namespace Palisades.Model
{
    [XmlInclude(typeof(LnkShortcut))]
    [XmlInclude(typeof(UrlShortcut))]
    public abstract class Shortcut : INotifyPropertyChanged
    {
        private string name;
        private string iconPath;
        private string uriOrFileAction;
        private string shellIconLocation;
        private string sourceShortcutPath;
        private string groupName;
        private string typeName;
        private bool isSelected;
        private bool isRenaming;
        private string pendingName;

        public Shortcut() : this("", "", "")
        {

        }
        public Shortcut(string name, string iconPath, string uriOrFileAction)
        {
            this.name = name;
            this.iconPath = iconPath;
            this.uriOrFileAction = uriOrFileAction;
            shellIconLocation = string.Empty;
            sourceShortcutPath = string.Empty;
            groupName = PalisadeModel.DefaultGroupName;
            typeName = string.Empty;
            isSelected = false;
            isRenaming = false;
            pendingName = name;
        }

        public string Name
        {
            get { return name; }
            set
            {
                name = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
                if (!IsRenaming)
                {
                    pendingName = name;
                }

                OnPropertyChanged();
            }
        }

        public string IconPath { get { return iconPath; } set { iconPath = value; OnPropertyChanged(); } }
        public string UriOrFileAction { get { return uriOrFileAction; } set { uriOrFileAction = value; OnPropertyChanged(); } }
        public string ShellIconLocation
        {
            get { return string.IsNullOrWhiteSpace(shellIconLocation) ? string.Empty : shellIconLocation; }
            set
            {
                shellIconLocation = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
                OnPropertyChanged();
            }
        }

        public string SourceShortcutPath
        {
            get { return string.IsNullOrWhiteSpace(sourceShortcutPath) ? string.Empty : sourceShortcutPath; }
            set
            {
                sourceShortcutPath = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
                OnPropertyChanged();
            }
        }
        public string GroupName
        {
            get { return string.IsNullOrWhiteSpace(groupName) ? PalisadeModel.DefaultGroupName : groupName; }
            set
            {
                groupName = string.IsNullOrWhiteSpace(value) ? PalisadeModel.DefaultGroupName : value.Trim();
                OnPropertyChanged();
            }
        }

        public string TypeName
        {
            get { return string.IsNullOrWhiteSpace(typeName) ? string.Empty : typeName; }
            set
            {
                typeName = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
                OnPropertyChanged();
            }
        }

        [XmlIgnore]
        public bool IsSelected
        {
            get { return isSelected; }
            set
            {
                isSelected = value;
                OnPropertyChanged();
            }
        }

        [XmlIgnore]
        public bool IsRenaming
        {
            get { return isRenaming; }
            set
            {
                isRenaming = value;
                OnPropertyChanged();
            }
        }

        [XmlIgnore]
        public string PendingName
        {
            get { return string.IsNullOrWhiteSpace(pendingName) ? Name : pendingName; }
            set
            {
                pendingName = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        /// FIXME: Would be cool to move it out the model.

        public static string GetName(string filename)
        {
            return Path.GetFileNameWithoutExtension(filename);
        }

        public static string GetIcon(string filename, string palisadeIdentifier)
        {
            using Bitmap icon = IconExtractor.GetFileImageFromPath(filename, Helpers.Native.IconSizeEnum.LargeIcon48);

            string iconDir = PDirectory.GetPalisadeIconsDirectory(palisadeIdentifier);
            PDirectory.EnsureExists(iconDir);

            string iconFilename = Guid.NewGuid().ToString() + ".png";
            string iconPath = Path.Combine(iconDir, iconFilename);
            using FileStream fileStream = new(iconPath, FileMode.Create);
            icon.Save(fileStream, ImageFormat.Png);

            return iconPath;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
