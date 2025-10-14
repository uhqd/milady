using System;
using System.IO;
using System.Windows;

namespace catchDose
{



    public class DocSettings
    {
        public String HostName { get; private set; }
        public String Port { get; private set; }
        public String DocKey { get; private set; }
        public String ImportDir { get; private set; }

        public static DocSettings ReadSettings()
        {


            DocSettings docSettings = new DocSettings();


            #region GET APIKEY
            // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            // Change path to a file with your credentials
            string filePath = @"\\srv015\radiotherapie\SCRIPTS_ECLIPSE_v18\apikey\v18.txt";
            // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            if (!File.Exists(filePath))
                MessageBox.Show("Please create a file with apikey, hostname and port. Cannot find:\n" + filePath);

            string[] lines = File.ReadAllLines(filePath);
            docSettings.DocKey = lines.Length > 0 ? lines[0] : string.Empty;
            docSettings.HostName = lines.Length > 1 ? lines[1] : string.Empty;
            docSettings.Port = lines.Length > 2 ? lines[2] : string.Empty;
            #endregion


            // seems useless
            //docSettings.ImportDir = @"\\srvaria15-img\va_data$\Documents";
            docSettings.ImportDir = @"\\srvaria18-platf\va_data$\Documents";



            return docSettings;
        }

        public static string RemoveWhitespace(string line)
        {
            return string.Join("", line.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries));
        }
    }

}
