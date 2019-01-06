using NetOffice.ExcelApi;
using NetOffice.OfficeApi;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Application = NetOffice.ExcelApi.Application;

namespace ExcelPythonAddIn.CustomXML
{
    class PythonSourceManager
    {
        private const string prefix = "PythonSource.";

        public static string GetSourceCode(Workbook wb, string sheetCodeName)
        {
            var v = CustomXml.Get(wb, prefix + sheetCodeName);
            return v;
        }
        

        public static void SetSourceCode(Workbook wb, string sheetCodeName, string sourceCode)
        {
            CustomXml.Set(wb, prefix + sheetCodeName, sourceCode);
            //todo: this is probably a good time to clean out source associated with any sheets that no longer exist
        }

        public static Dictionary<string,string> GetAllSourceCode(Workbook wb)
        {
            return CustomXml.GetAll(wb, prefix);
        }
    }
}
