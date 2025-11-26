using catchDose;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;


// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(VMS.TPS.Common.Model.API.ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            #region Set current dir to the dll dir
            string fullPath = Assembly.GetExecutingAssembly().Location; //get the full location of the assembly            
            string theDirectory = System.IO.Path.GetDirectoryName(fullPath);//get the folder that's in                                                                  
            Directory.SetCurrentDirectory(theDirectory);// set current directory as the .dll directory...
            #endregion

            #region Check if a patient  is loaded


            try
            {
                string s = context.Patient.Id; // check if a patient is loaded
            }
            catch
            {
                System.Windows.MessageBox.Show("Merci de charger un patient");
                return;
            }

            
            #endregion


            PreliminaryInformation pinfo = new PreliminaryInformation(context);
            /*
            _ = typeof(UglyToad.PdfPig.PdfDocument);          // PdfSharp
             _ = typeof(JsonConvert);          // Newtonsoft.Json


            var assemblies = AppDomain.CurrentDomain.GetAssemblies().OrderBy(a => a.GetName().Name);
            String msg = string.Empty;
            foreach (var asm in assemblies)
            {
                AssemblyName name = asm.GetName();
                msg += name.Name + " " + name.Version + "\n";
            }
            File.AppendAllText(@"\\srv015\sf_com\simon_lu\assemblies.txt", msg);
            */
        }
    }
}
