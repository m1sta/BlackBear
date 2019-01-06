using ExcelPythonAddIn.CustomXML;
using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPythonAddIn.Ribbon
{
    class RefreshMulti
    {
        public static void GetAfter(Workbook wb, string sheetName)
        {
            //todo: consider tables names as well as sheet names
            CTPManager.Save();
            var sourceDict = PythonSourceManager.GetAllSourceCode(wb).Where(i => i.Value.Trim().Length > 0);
            var sheetNames = wb.Worksheets.Select(s => (s as Worksheet).Name);
            var proposed = new HashSet<string> { sheetName };
            var parsed = new Dictionary<string, List<string>> { };
            var executed = new HashSet<string> { };

            if (sourceDict.ToList().Count == 0)
            {
                MessageBox.Show("This workbook does not contain any embedded python code.");
                return;
            }

            while (proposed.Where(prop => !parsed.Keys.Contains(prop)).ToList().Count > 0)
            {
                var toParse = proposed.Where(prop => !parsed.Keys.Contains(prop));
                var proposedSheet = toParse.First();
                var source = sourceDict.Where(x => x.Key == proposedSheet).First().Value;
                var successors = sourceDict.Where(s=>proposed.Where(p=>s.Value.Contains(p)).ToList().Count>0);
                var predecessors = proposed.Where(sh => source.Contains(sh) && sourceDict.Where(i => i.Key == sh).ToList().Count > 0).ToList();

                foreach (var s in successors) proposed.Add(s.Key);
                parsed.Add(proposedSheet, predecessors);
            }

            executed.Add(sheetName);
            RefreshTable.Start(sheetName);

            while (parsed.Keys.Where(k => !executed.Contains(k)).ToList().Count > 0)
            {
                var notExecuted = parsed.Where(k => !executed.Contains(k.Key));
                var readyToExecute = notExecuted.Where(i => i.Value.Where(p => !executed.Contains(p)).ToList().Count == 0);
                if (readyToExecute.ToList().Count == 0 && notExecuted.ToList().Count > 0)
                {
                    MessageBox.Show("There is a circular reference within your scripts. The following scripts were not executed: \n" + string.Join(", ", notExecuted.Select(i=>i.Key)));
                    break;
                }
                else
                {
                    var selected = readyToExecute.First().Key;
                    RefreshTable.Start(selected);
                    executed.Add(selected);
                }

            }
        }

        public static void GetBefore(Workbook wb, string sheetName)
        {
            //todo: consider tables names as well as sheet names
            CTPManager.Save();
            var sourceDict = PythonSourceManager.GetAllSourceCode(wb).Where(i => i.Value.Trim().Length > 0);
            var sheetNames = wb.Worksheets.Select(s => (s as Worksheet).Name);
            var proposed = new HashSet<string> { sheetName };
            var parsed = new Dictionary<string, List<string>> { };
            var executed = new HashSet<string> { };

            if(sourceDict.ToList().Count == 0)
            {
                MessageBox.Show("This workbook does not contain any embedded python code.");
                return;
            }

            while(proposed.Where(prop=>!parsed.Keys.Contains(prop)).ToList().Count > 0)
            {
                var toParse = proposed.Where(prop => !parsed.Keys.Contains(prop));
                var proposedSheet = toParse.First();
                var source = sourceDict.Where(x => x.Key == proposedSheet).First().Value;
                var predecessors = sheetNames.Where(sh => source.Contains(sh) && sourceDict.Where(i => i.Key == sh).ToList().Count > 0).ToList();
                foreach (var p in predecessors) proposed.Add(p);
                parsed.Add(proposedSheet, predecessors);
            }

            while (parsed.Keys.Where(k => !executed.Contains(k)).ToList().Count > 0)
            {
                var notExecuted = parsed.Where(k => !executed.Contains(k.Key));
                var readyToExecute = notExecuted.Where(i => i.Value.Where(p => !executed.Contains(p)).ToList().Count == 0);
                if(readyToExecute.ToList().Count==0 && notExecuted.ToList().Count > 0)
                {
                    MessageBox.Show("There is a circular reference within your scripts. The following scripts were not executed: \n" + string.Join(", ", notExecuted));
                    break;
                } else
                {
                    var selected = readyToExecute.First().Key;
                    RefreshTable.Start(selected);
                    executed.Add(selected);
                }
                
            }
        }
    }
}
