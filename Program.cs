using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelTools
{
    static class Program
    {
        [DllImport("kernel32.dll")]
        public static extern bool AllocConsole();


        [DllImport("kernel32.dll")]
        static extern bool FreeConsole();


        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AllocConsole();//调用系统API，调用控制台窗口

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new Form1());
            FreeConsole();//释放控制台
        }
    }
}
