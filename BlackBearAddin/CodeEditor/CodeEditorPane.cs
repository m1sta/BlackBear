
using ExcelDna.Integration.CustomUI;
using Application = NetOffice.ExcelApi.Application;
using NetOffice.ExcelApi;
using NetOffice;
using ExcelPythonAddIn.CustomXML;
using System;

namespace ExcelPythonAddIn
{
    public class CodeEditorPane
    {
        private CustomTaskPane _ctp;
        public CodeEditor _ce;
        private Application _excel;

        public CodeEditorPane(CustomTaskPane ctp, CodeEditor ce){
            _ctp = ctp;
            _ce = ce;
            _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);

            _ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            _ctp.VisibleStateChange += _ctp_VisibleStateChange;
            _ce.Size = new System.Drawing.Size(_ctp.Width, _ctp.Height);
            _ctp.Visible = true;

            _excel.ActiveWorkbook.BeforeSaveEvent += ActiveWorkbook_BeforeSaveEvent;
            _excel.ActiveWorkbook.SheetActivateEvent += _excel_SheetActivateEvent;
            _excel.ActiveWorkbook.SheetDeactivateEvent += _excel_SheetDeactivateEvent;

            //Load existing source for current sheet
            var shObj = _excel.ActiveSheet as Worksheet;
            _ce.SourceCode = PythonSourceManager.GetSourceCode(shObj.Parent as Workbook, shObj.Name);

        }

        #region EventHandlers

        private void _excel_SheetDeactivateEvent(COMObject COMSh)
        {
            var Sh = COMSh as Worksheet;
            this.Save(Sh);
        }

        internal void Load()
        {
            var shObj = _excel.ActiveSheet as Worksheet;
            _ce.SourceCode = PythonSourceManager.GetSourceCode(shObj.Parent as Workbook, shObj.Name);
        }

        private void _excel_SheetActivateEvent(COMObject Sh)
        {
            var shObj = Sh as Worksheet;
            var retrieved = PythonSourceManager.GetSourceCode(shObj.Parent as Workbook, shObj.Name);
            _ce.SourceCode = retrieved;
        }

        private void ActiveWorkbook_BeforeSaveEvent(bool SaveAsUI, ref bool Cancel)
        {
            this.Save();
        }

        private void _ctp_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {          
            this.Save();
        }

        #endregion

        public void Save(Worksheet targetSheet = null, string newSource = null)
        {
            if (targetSheet == null) targetSheet = _excel.ActiveSheet as Worksheet;
            if(newSource != null) _ce.SourceCode = newSource;
            if (_ce.SourceCode != "!") PythonSourceManager.SetSourceCode(targetSheet.Parent as Workbook, targetSheet.Name, _ce.SourceCode);
        }

        public void Open()
        {
            _ctp.Visible = true;
        }

        public void Close()
        {
            _ctp.Visible = false;
        }

        public bool Opened
        {
            get
            {
                return _ctp.Visible;
            }
        }
    }
}