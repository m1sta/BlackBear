using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using Application = NetOffice.ExcelApi.Application;
using NetOffice.ExcelApi;
using System.Threading;
using System.Collections.Generic;

namespace ExcelPythonAddIn
{
    public static class CTPManager
    {
        private static Dictionary<int, CodeEditorPane> ctpList = new Dictionary<int, CodeEditorPane>();
        private static Application _excel;

        public static void Show()
        {
            _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            if (ctpList.ContainsKey(Parent) == false)
            {
                CustomTaskPane ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(CodeEditor), "Python Code Editor");
                CodeEditorPane cep = new CodeEditorPane(ctp, (CodeEditor) ctp.ContentControl);
                _excel.WorkbookActivateEvent += _excel_WorkbookActivateEvent;

                ctpList.Add(Parent, cep);
            }
            else {
                if (ctpList[Parent].Opened) {
                    Save();
                }
            }
            ctpList[Parent].Open();

        }

        private static void _excel_WorkbookActivateEvent(Workbook Wb)
        {
            foreach (KeyValuePair<int, CodeEditorPane> ctp in ctpList)
            {
                ctp.Value.Load();
            }
        }

        public static void CloseCTP()
        {
            if (ctpList.ContainsKey(Parent) == true)
            {
                ctpList[Parent].Close();
            }
        }

        public static void Save(string newSource = null)
        {
            if (ctpList.ContainsKey(Parent) == true)
            {
                ctpList[Parent].Save(newSource:newSource);
            }
        }
        
        private static int Parent
        {
            get
            {
                _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
                return _excel.ActiveWindow.Hwnd;
            }
        }
         
    }
}
