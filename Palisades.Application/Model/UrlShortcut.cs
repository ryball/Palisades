using System.IO;
using System.Linq;

namespace Palisades.Model
{
    public class UrlShortcut : Shortcut
    {

        public UrlShortcut() : base()
        {
        }
        public UrlShortcut(string name, string iconPath, string uriOrFileAction) : base(name, iconPath, uriOrFileAction)
        {
        }

        public static UrlShortcut? BuildFrom(string shortcut, string palisadeIdentifier)
        {
            string[] lines = File.ReadAllLines(shortcut);
            string? line = lines.FirstOrDefault((value) => value.StartsWith("URL="));
            if (line == null)
            {
                return null;
            }

            string url = line.Replace("URL=", "");
            url = url.Replace("\"", "");
            url = url.Replace("BASE", "");

            string? iconFile = lines.FirstOrDefault(value => value.StartsWith("IconFile="))?.Replace("IconFile=", "").Trim();
            string? iconIndex = lines.FirstOrDefault(value => value.StartsWith("IconIndex="))?.Replace("IconIndex=", "").Trim();
            string shellIconLocation = string.IsNullOrWhiteSpace(iconFile)
                ? string.Empty
                : (int.TryParse(iconIndex, out int parsedIndex) ? $"{iconFile},{parsedIndex}" : iconFile);

            string name = Shortcut.GetName(shortcut);
            string iconPath = Shortcut.GetIcon(shortcut, palisadeIdentifier);

            return new UrlShortcut(name, iconPath, url)
            {
                ShellIconLocation = shellIconLocation,
                SourceShortcutPath = shortcut
            };
        }
    }
}
