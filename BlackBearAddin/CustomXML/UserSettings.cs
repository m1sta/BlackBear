using NetOffice.ExcelApi;
using System; 
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Application = NetOffice.ExcelApi.Application;

namespace ExcelPythonAddIn.CustomXML
{
    public class UserSettings
    {
        private Workbook _wb;

        public UserSettings(Workbook wb)
        {
            _wb = wb;
        }

        public string Get(string prop)
        {
            return CustomXml.Get(_wb, prop);
        }

        public void Set(string prop, string val)
        {
            CustomXml.Set(_wb, prop, val);
        }
        public string path
        {
            get {
                return this.Get("PythonPath").Trim().Length > 0 ? Get("PythonPath") : "python";
            }
            set {
                this.Set("PythonPath", value);
            }
        }

        public string args
        {
            get
            {
                return this.Get("PythonExecutionString").Trim().Length > 0 ? Get("PythonExecutionString") : "";
            }
            set
            {
                this.Set("PythonExecutionString", value);
            }
        }

        public string prepend
        {
            get
            {
                var defaultPrependPath = Path.Combine(Path.GetDirectoryName(ExcelDna.Integration.ExcelDnaUtil.XllPath), "DefaultPrepend.py");
                var defaultVal = File.ReadAllText(defaultPrependPath);
                var val = this.Get("PythonPrependedCode").Trim().Length > 0 ? Get("PythonPrependedCode") : defaultVal;
                return val;
            }
            set
            {
                this.Set("PythonPrependedCode", value);
            }
        }
    }
}
