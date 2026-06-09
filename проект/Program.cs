using System;
using System.Windows.Forms;

namespace BytovkaBuilder
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());  // Здесь должно быть Form1()
        }
    }
}