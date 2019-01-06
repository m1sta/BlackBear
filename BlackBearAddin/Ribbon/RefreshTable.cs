using ExcelPythonAddIn.CustomXML;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = NetOffice.ExcelApi.Application;

namespace ExcelPythonAddIn
{
    class RefreshTable
    {
        
        public static void Start(string sheetName="")
        {

            CTPManager.Save();
            var _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);

            var activeSheet = sheetName == "" ? (Worksheet)_excel.ActiveSheet : (Worksheet) _excel.ActiveWorkbook.Worksheets[sheetName];
            string sourceCode = PythonSourceManager.GetSourceCode(activeSheet.Parent as Workbook, activeSheet.Name);

            //remove non-printable characters
            sourceCode = Regex.Replace(sourceCode, @"[^ -~\n\t]+", " ");

            if (sourceCode.Length > 0)
            {
                var settings = new CustomXML.UserSettings(_excel.ActiveWorkbook);
                var workbookPath = "\"" + _excel.ActiveWorkbook.Path.Replace("\\", "\\\\") + "\"";
                var expandedSourceCode = "workbookPath=" + workbookPath 
                    + "\nimport os, pandas\nif(len(workbookPath)>1): os.chdir(workbookPath)\nresult = pandas.DataFrame()\n" 
                    + settings.prepend + "\n" 
                    + sourceCode 
                    + "\nif(result.shape[0] != 0): table[this] = result";

                var tempCol = new TempFileCollection(Path.GetTempPath(), false);
                var filename = tempCol.AddExtension("py");
                File.WriteAllText(filename, expandedSourceCode);

                //Run Python code
                _excel.StatusBar = $"Running Python for '{activeSheet.Name}' In Background ...";
                Run(settings.path, Path.Combine(settings.args, filename), sourceCode, activeSheet);
                _excel.StatusBar = "";
                //
            }
            else
            {
                MessageBox.Show("There is no table refresh code associated with this worksheet.", "Refresh Table Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private static void Run(string pyPath, string pyArg, string sourceCode, Worksheet ws)
        {
            var _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            try
            {
                var psi = new ProcessStartInfo()
                {
                    FileName = pyPath,
                    WorkingDirectory = Path.GetDirectoryName(_excel.ActiveWorkbook.Path.Length > 0 ? _excel.ActiveWorkbook.Path : pyArg),
                    Arguments = pyArg,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                var process = new Process();
                var output = "";
                var error = "";
                process.StartInfo = psi;
                process.EnableRaisingEvents = false;
                process.OutputDataReceived += (sender, eventArgs) => output += eventArgs.Data + "\n";
                process.ErrorDataReceived += (sender, eventArgs) => error += eventArgs.Data + "\n";
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                var processExited = process.WaitForExit(30000);
                var lastError = error.Trim().Split('\n').Reverse().Take(1).ToArray()[0].ToString();

                if (processExited == false) // we timed out...
                {
                    process.Kill();
                    lastError = "Execution of the python script timed out after 30 seconds.";
                }

                //Script execution completed.
                var separator = new string[] { "\n\n#Refreshed = ", "\n \n#Refreshed = " };
                var newSource = sourceCode.Split(separator, StringSplitOptions.None).FirstOrDefault();
                newSource += "\n\n#Refreshed = " + DateTime.Now.ToString("yyyy-mm-dd @ hh:mm tt");
                if (lastError.Trim().Length > 0) newSource += "\n#Error = " + lastError;
                if (error.Trim().Length > 0 || true) newSource += "\n#Source = " + pyArg;
                if (output.Trim().Length > 0) newSource += "\n#Output = " + output.Trim().Replace("\n", "\n#Output = ");
                if (error.Trim().Length > 0) newSource += "\n\n#StackTrace\n#" + error.Replace("\n", "\n#");
                PythonSourceManager.SetSourceCode((Workbook) ws.Parent, ws.Name, newSource);
                if (ws.Name == ((ws.Parent as Workbook).ActiveSheet as Worksheet).Name) CTPManager.Save(newSource);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error executing Python script: " + ex.Message);
            }
            finally
            {
                _excel.Application.EnableEvents = true;
                _excel.Application.ScreenUpdating = true;
            }
        }
    }
}
