using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using VB2C;

namespace VisualBasicUpgradeAssistant.WinForms
{
    sealed class Program
    {
        public static XMLConfig Config;
        public static frmConvert MainForm;

        // configuration constants
        private const string CONFIG_FILE = "vb2c.xml";
        public const string CONFIG_SETTING = "Setting";
        public const string CONFIG_OUT_PATH = "OutPath";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // get current directory
            string[] CommandLineArgs;
            CommandLineArgs = Environment.GetCommandLineArgs();
            // index 0 contain path and name of exe file
            string sBinPath = Path.GetDirectoryName(CommandLineArgs[0].ToLower());

            // create configuration object
            Config = new XMLConfig(sBinPath + Path.DirectorySeparatorChar + CONFIG_FILE);

            // create main screen
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmConvert());
        }
    }
}
