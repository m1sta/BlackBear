using System; 
using System.Runtime.InteropServices; 
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using System.Drawing;
using System.Windows.Forms;
using ExcelPythonAddIn.Ribbon;
using NetOffice.ExcelApi;

namespace ExcelPythonAddIn
{
    [ComVisible(true)]
    public class TableStreamRibbon : ExcelRibbon
    {
        private Application _excel;
        private IRibbonUI _thisRibbon;

        public override string GetCustomUI(string ribbonId)
        {
            _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            _excel.WorkbookActivateEvent += _excel_WorkbookActivateEvent;
            return Properties.Resources.TableStreamRibbon;       
        }

        private void _excel_WorkbookActivateEvent(NetOffice.ExcelApi.Workbook Wb)
        {
            var refTable = new RefreshTable();
        }
         

        public void OnLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));
            }

            _thisRibbon = ribbon;

            _excel.WorkbookActivateEvent += OnInvalidateRibbon;
            _excel.WorkbookDeactivateEvent += OnInvalidateRibbon;
            _excel.SheetActivateEvent += OnInvalidateRibbon;
            _excel.SheetDeactivateEvent += OnInvalidateRibbon;

            if (_excel.ActiveWorkbook == null)
            { 
                _excel.Workbooks.Add();
            } 

        }

        private void OnInvalidateRibbon(object obj)
        {
            _thisRibbon.Invalidate();
        }

        public void OnPressMe(IRibbonControl control)
        {
            using (var controller = new ExcelController(_excel, _thisRibbon))
            {
                switch (control.Id)
                {
                    case "btnShowEditor":
                        controller.btnShowEditor_Click();
                        break;  
                    case "btnSettings": 
                        controller.btnSettings_Click();
                        break;
                    case "btnRefreshTable":
                        controller.btnRefreshTable_Click();
                        break;
                    case "btnViewBefore":
                        RefreshMulti.GetBefore(_excel.ActiveWorkbook, (_excel.ActiveSheet as Worksheet).Name);
                        break;
                    case "btnViewAfter":
                        RefreshMulti.GetAfter(_excel.ActiveWorkbook, (_excel.ActiveSheet as Worksheet).Name);
                        //controller.btnViewDependencies_Click();
                        break;
                    default:
                        break;
                }
            }
        }
        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnShowEditor": return new Bitmap(Properties.Resources.ShowEditor);
                case "btnSettings": return new Bitmap(Properties.Resources.Settings);
                case "btnRefreshTable": return new Bitmap(Properties.Resources.RefreshTable);
                case "btnViewDependencies": return new Bitmap(Properties.Resources.RefreshTable);
                default: return null;
            }
        }
    }
}
