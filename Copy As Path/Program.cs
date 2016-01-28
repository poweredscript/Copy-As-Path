using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IWshRuntimeLibrary;
using Microsoft.Win32;

namespace Copy_As_Path
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (IsRunAsAdministrator())
            {
                RegistyCheck();
                return;
            }
            //StringCollection paths = new StringCollection();
            //paths.Add("f:\\temp\\test.txt");
            //paths.Add("f:\\temp\\test2.txt");
            //Clipboard.SetFileDropList(paths);

            if (args != null && args.Length > 0)
            {
                for (int i = 0; i < 3; i++)
                {
                    Clipboard.SetText(string.Join("\n", args));
                    var text = Clipboard.GetText(TextDataFormat.Text);
                    if (!string.IsNullOrEmpty(text))
                    {
                        Console.Beep();
                        break;
                    }
                }
               
            }
            //Console.Write(string.Join("\n", args));
            //Console.ReadKey();
        }
        public static void RegistyCheck()
        {
            var fullName = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var thisPath = Directory.GetCurrentDirectory();
            //var iconPath = Path.Combine(thisPath, "Note.ico");
            var iconPathReg = "\"" + fullName + "\"" + "";
            var exePathReg = "\"" + fullName + "\"" + " \"%1\"";
            var exePathReg2 = "\"" + fullName + "\"" + " \"%V\"";
            //File
            Registry.SetValue(@"HKEY_CLASSES_ROOT\*\shell\Copy File Path", "Icon", iconPathReg, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CLASSES_ROOT\*\shell\Copy File Path\command", "", exePathReg, RegistryValueKind.String);
            //Folder
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\shell\Copy Folder Path", "Icon", iconPathReg, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\shell\Copy Folder Path\command", "", exePathReg, RegistryValueKind.String);
            //Inside folder
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\Background\shell\Copy Folder Path", "Icon", iconPathReg, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\Background\shell\Copy Folder Path\command", "", exePathReg2, RegistryValueKind.String);

            var schPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Windows\SendTo\Copy As Path.lnk";
            //MessageBox.Show(fullName);
            //MessageBox.Show(thisPath);
            //MessageBox.Show(schPath);
            //MessageBox.Show(fullName);
            if (System.IO.File.Exists(schPath)) System.IO.File.Delete(schPath);
            WshShellClass wsh = new WshShellClass();
            IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(schPath) as IWshRuntimeLibrary.IWshShortcut;
            shortcut.Arguments = "";
            shortcut.TargetPath = fullName;
            // not sure about what this is for
            shortcut.WindowStyle = 1;
            shortcut.Description = "Copy As Path";
            shortcut.WorkingDirectory = thisPath;
            shortcut.IconLocation = fullName;
            shortcut.Save();
            MessageBox.Show("Success Install", "Success Install", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private static bool IsRunAsAdministrator()
        {
            WindowsIdentity wi = WindowsIdentity.GetCurrent();
            WindowsPrincipal wp = new WindowsPrincipal(wi);

            return wp.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
