using System;
using System.Runtime.InteropServices;

namespace Palisades.Model
{
    public class LnkShortcut : Shortcut
    {

        public LnkShortcut() : base()
        {
        }
        public LnkShortcut(string name, string iconPath, string uriOrFileAction) : base(name, iconPath, uriOrFileAction)
        {
        }

        public static LnkShortcut? BuildFrom(string shortcut, string palisadeIdentifier)
        {
            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                return null;
            }

            object? shell = null;
            object? link = null;

            try
            {
                shell = Activator.CreateInstance(shellType);
                if (shell == null)
                {
                    return null;
                }

                link = shellType.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { shortcut });
                if (link == null)
                {
                    return null;
                }

                string? targetPath = link.GetType().InvokeMember("TargetPath", System.Reflection.BindingFlags.GetProperty, null, link, null) as string;
                if (string.IsNullOrWhiteSpace(targetPath))
                {
                    return null;
                }

                string name = Shortcut.GetName(shortcut);
                string iconPath = Shortcut.GetIcon(shortcut, palisadeIdentifier);

                return new LnkShortcut(name, iconPath, targetPath);
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
    }
}
