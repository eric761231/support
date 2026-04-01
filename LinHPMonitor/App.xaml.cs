using System;
using System.IO;
using System.Windows;

namespace LinHPMonitor
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += (s, args) => {
                File.WriteAllText("fatal_crash.log", args.ExceptionObject.ToString());
                MessageBox.Show("Fatal Error: Check fatal_crash.log\n" + args.ExceptionObject.ToString());
            };
            base.OnStartup(e);
        }
    }
}
