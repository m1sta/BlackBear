using ExcelDna.Integration; 
using System.IO; 

namespace ExcelPythonAddIn
{
    public class AssemblyInfo
    {

        public static string AssemblyPath
        {
            get
            {
                return Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            }
        }
    }
}
