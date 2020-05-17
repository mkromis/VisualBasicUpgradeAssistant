using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using VB2C;

namespace VisualBasicUpgradeAssistant.WinForms
{
    public sealed class Program
    {
        public static XMLConfig Config;
        public static FrmConvert MainForm;

        // configuration constants
        private const String CONFIG_FILE = "vb2c.xml";
        public const String CONFIG_SETTING = "Setting";
        public const String CONFIG_OUT_PATH = "OutPath";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            // get current directory
            String[] CommandLineArgs;
            CommandLineArgs = Environment.GetCommandLineArgs();
            // index 0 contain path and name of exe file
            String sBinPath = Path.GetDirectoryName(CommandLineArgs[0].ToLower());

            // create configuration object
            Config = new XMLConfig(sBinPath + Path.DirectorySeparatorChar + CONFIG_FILE);

            // create main screen
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmConvert());
        }
    }
}
