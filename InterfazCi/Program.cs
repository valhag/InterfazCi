using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace InterfazCi
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            //Application.Run(new Gomar());
            //Application.Run(new Generatxt());
            //Application.Run(new Romero());
            //Application.Run(new SapToCont());
            Application.Run(new TabPolizas());
        }
    }
}
