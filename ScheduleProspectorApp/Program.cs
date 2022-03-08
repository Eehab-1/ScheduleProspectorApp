using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScheduleProspectorApp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

           Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("NTkyNTc1QDMxMzkyZTM0MmUzMEZYbFByN21GZlZGRDNFTVZiRTZCQkxxdGl3MCtwYTBlSHdnTjRlbVZuaHM9");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MetroForm1());
        }
    }
}
