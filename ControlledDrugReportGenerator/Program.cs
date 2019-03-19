using ControlledDrugReportGenerator.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ControlledDrugReportGenerator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool isNotDuplicateApplication = false;
            using (Mutex mutex = new Mutex(true, Application.ProductName, out isNotDuplicateApplication))
            {
                if (isNotDuplicateApplication)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new FormMain());
                }
                else
                {
                    MessageBox.Show("請不要重複開啟管制藥報表程式", "Error");
                    Application.Exit();
                }
            }

        }
    }
}
