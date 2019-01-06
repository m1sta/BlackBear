using Microsoft.Toolkit.Win32.UI.Controls.WinForms;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelPythonAddIn
{
    /////////////// Define the UserControl to display on the CTP ///////////////////////////
    // Would need to be marked with [ComVisible(true)] if in a project that is marked as [assembly:ComVisible(false)] which is the default for VS projects.
    [ComVisible(true)]
    public class CodeEditor : UserControl
    {
        private WebView webView1;
        private SimpleHTTPServer monacoServer;
        private string _pendingSourceUpdate = "!";

        public CodeEditor()
        {
            InitializeComponent();
            var monacoFolder = Path.Combine(Path.GetDirectoryName(ExcelDna.Integration.ExcelDnaUtil.XllPath), "Monaco");
            monacoServer = new SimpleHTTPServer(monacoFolder);
            var monacoUrl = "http://localhost:" + monacoServer.Port.ToString();
            webView1.Source = new Uri(monacoUrl);

            webView1.NavigationStarting += WebView1_NavigationStarting;
            webView1.NavigationCompleted += WebView1_NavigationCompleted;
            webView1.ScriptNotify += WebView1_ScriptNotify;
         }

        private void WebView1_ScriptNotify(object sender, Microsoft.Toolkit.Win32.UI.Controls.Interop.WinRT.WebViewControlScriptNotifyEventArgs e)
        {
            Debug.Print("Script Notify: " + e.Value.ToString());
            RefreshTable.Start();
        }

        private void WebView1_NavigationCompleted(object sender, Microsoft.Toolkit.Win32.UI.Controls.Interop.WinRT.WebViewControlNavigationCompletedEventArgs e)
        {
            System.Diagnostics.Debug.Print("Completed: " + webView1.Source.ToString());
            if (_pendingSourceUpdate != "!")
            {
                this.SourceCode = _pendingSourceUpdate;
                _pendingSourceUpdate = "!";
            }
        }

        private void WebView1_NavigationStarting(object sender, Microsoft.Toolkit.Win32.UI.Controls.Interop.WinRT.WebViewControlNavigationStartingEventArgs e)
        {
            System.Diagnostics.Debug.Print("Starting: " + e.Uri.ToString());
        }

        public string SourceCode
        {
            get
            {
                try
                {
                    var source = webView1.InvokeScript("getSource");
                    return source;
                }
                catch (Exception ex)
                {
                    return "!";
                }
                
            }
            set
            {
                try
                {
                    webView1.InvokeScript("setSource", new string[] { value });
                } 
                catch (Exception ex)
                {
                    _pendingSourceUpdate = value;
                }
            }
        }
        private void InitializeComponent()
        {
            this.webView1 = new Microsoft.Toolkit.Win32.UI.Controls.WinForms.WebView();
            ((System.ComponentModel.ISupportInitialize)(this.webView1)).BeginInit();
            this.SuspendLayout();
            // 
            // webView1
            // 
            this.webView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webView1.IsPrivateNetworkClientServerCapabilityEnabled = true;
            this.webView1.IsScriptNotifyAllowed = true;
            this.webView1.Location = new System.Drawing.Point(0, 0);
            this.webView1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webView1.Name = "webView1";
            this.webView1.Size = new System.Drawing.Size(681, 306);
            this.webView1.TabIndex = 0;
            // 
            // CodeEditor
            // 
            this.Controls.Add(this.webView1);
            this.Name = "CodeEditor";
            this.Size = new System.Drawing.Size(681, 306);
            ((System.ComponentModel.ISupportInitialize)(this.webView1)).EndInit();
            this.ResumeLayout(false);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var source = webView1.InvokeScript("getSource");
            MessageBox.Show(source.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var source = webView1.InvokeScript("setSource", new string[] {"test"});
            MessageBox.Show(source.ToString());
        }
    }
}
