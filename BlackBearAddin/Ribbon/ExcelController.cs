using System;
using System.Linq;
using System.Windows;
using ExcelDna.Integration.CustomUI;
using ExcelPythonAddIn.CustomXML;
using NetOffice.ExcelApi;
using Application = NetOffice.ExcelApi.Application;

namespace ExcelPythonAddIn
{
    class ExcelController : IDisposable
    {
        private readonly IRibbonUI _modelingRibbon;
        protected readonly Application _excel;

        public ExcelController(Application excel, IRibbonUI modelingRibbon)
        {
            _modelingRibbon = modelingRibbon;
            _excel = excel;
        }

        public void PressMe(string id)
        { 
            var activeSheet = _excel.ActiveSheet as Worksheet;
            activeSheet.Range("A1").Value = $"{id}, {id.ToLower()}, {id.ToLower()}";
        }

        public void btnShowEditor_Click()
        {
            CTPManager.Show(); 
        }
                         
        public void btnSettings_Click() {
            var settings = new CustomXML.UserSettings(_excel.ActiveWorkbook);

            SettingsForm sf = new SettingsForm(_excel);
            
            sf.txtPyPath.Text = settings.path;
            sf.txtExeString.Text = settings.args;
            sf.txtPrepCode.Text = settings.prepend;

            sf.ShowDialog();
            
        }

        public void btnRefreshTable_Click()
        {
            RefreshTable.Start();
        }
         
        public void Dispose()
        {
        }

        internal void btnViewDependencies_Click()
        {
            //todo: display a new pane which visualises the dependencies as a graph and allows the scripts to be run by interacting with that graph
            var sourceDict = PythonSourceManager.GetAllSourceCode(_excel.ActiveWorkbook).Where(i=>i.Value.Trim().Length>0);
            var sheetNames = _excel.ActiveWorkbook.Worksheets.Select(s => (s as Worksheet).Name);
            foreach(var pair in sourceDict)
            {
                var predecessors = sheetNames.Where(sheet => pair.Value.Contains(sheet)).Select(n=>"'"+n+"'");
                if (predecessors.Count(i => true) == 0) MessageBox.Show("The worksheet '" + pair.Key + "' has no dependencies.");
                else
                {
                    var message = "The worksheet '" + pair.Key + "' uses data from the worksheet" + (predecessors.Count(i=>true)==1 ? " " : "s ") + String.Join(", ", predecessors) + ".";
                    MessageBox.Show(ReplaceLastOccurrence(message, ", ", " and "));
                }
            }
            if (sourceDict.Count(i => true) == 0) MessageBox.Show("There is no Python code embedded in this workbook.");

            //
        }

        public static string ReplaceLastOccurrence(string Source, string Find, string Replace)
        {
            int Place = Source.LastIndexOf(Find);
            string result = Place < 0 ? Source : Source.Remove(Place, Find.Length).Insert(Place, Replace);
            return result;
        }
    }
}
