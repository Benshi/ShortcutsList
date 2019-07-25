using System;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Threading;

// Install-Package Codeer.Friendly.Windows -Version 2.13.1
using Codeer.Friendly;
using Codeer.Friendly.Dynamic;
using Codeer.Friendly.Windows;

// Install-Package VSHTC.Friendly.PinInterface -Version 1.3.3
using VSHTC.Friendly.PinInterface;

// 参照設定 - COM - "Microsoft Development Environment 8.0"

namespace Orator.VSShortcut
{
    public class Program
    {
        public static void Main()
        {
            string outputFile = @"D:\shortcuts.txt";
            Process proc;

#if VSを起動してアタッチ
            string vs2017 = @"C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\devenv.exe";
            var proc = Process.Start(vs2017);
            proc.WaitForInputIdle();
            while (string.IsNullOrEmpty(proc.MainWindowTitle))
            {
                Thread.Sleep(10);
                proc = Process.GetProcessById(proc.Id);
            }
#else
            // 現在デバッグ実行している VS にアタッチ
            proc = Process.GetProcessesByName("devenv").FirstOrDefault();
#endif
            WindowsAppFriend app = new WindowsAppFriend(proc);
            app.LoadAssembly(typeof(Program).Assembly);

            // https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.package.getglobalservice
            AppVar obj = app.Type().Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(EnvDTE._DTE));

            using (var sw = File.CreateText(outputFile))
            {
                sw.WriteLine("Guid,ID,Name,LocalizedName,#,Shortcut");
                Console.Clear();

                // https://docs.microsoft.com/en-us/dotnet/api/envdte80.dte2
                // https://docs.microsoft.com/en-us/dotnet/api/envdte80.dte2.commands
                EnvDTE80.DTE2 dte2 = obj.Pin<EnvDTE80.DTE2>();
                EnvDTE.Commands commands = dte2.Commands;

                int count = commands.Count;
                for (int i = 1; i <= count; i++)
                {
                    Console.SetCursorPosition(0, 0);
                    Console.Write($"進捗 {i,7:N0} / {count,-7:N0}");

                    // https://docs.microsoft.com/en-us/dotnet/api/envdte.commands.item
                    // https://docs.microsoft.com/en-us/dotnet/api/envdte.command
                    EnvDTE.Command cmd = commands.Item(i, -1);

                    if (string.IsNullOrEmpty(cmd.Name)) { continue; }
                    if (string.IsNullOrEmpty(cmd.LocalizedName)) { continue; }

                    // https://docs.microsoft.com/en-us/dotnet/api/envdte.command.bindings
                    int b = 0;
                    foreach (string bind in cmd.Bindings)
                    {
                        sw.WriteLine($"{cmd.Guid}\t{cmd.ID}\t{cmd.Name}\t{cmd.LocalizedName}\t{++b}\t{bind}");
                    }
                }
            }
            Console.WriteLine("列挙完了");
            Process.Start(outputFile);
        }
    }
}
