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
            AllocConsole();//����ϵͳAPI�����ÿ���̨����

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new Form1());
            FreeConsole();//�ͷſ���̨
        }
    }
}
