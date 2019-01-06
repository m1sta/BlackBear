using System;
using System.Windows.Forms;
using Application = NetOffice.ExcelApi.Application;

namespace ExcelPythonAddIn
{
    public partial class SettingsForm : Form
    {

        public SettingsForm(Application xl)
        {
            _excel = xl;
            InitializeComponent();
        }

        protected readonly Application _excel;

        ///<summary>Default instance for static referencing.</summary>
        public static SettingsForm Default { get; private set; }

        private void btnSave_Click(object sender, EventArgs e)
        {

            var settings = new CustomXML.UserSettings(_excel.ActiveWorkbook);
            settings.path = txtPyPath.Text;
            settings.args = txtExeString.Text;
            settings.prepend = txtPrepCode.Text;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

            this.Close();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {

        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            var settings = new CustomXML.UserSettings(_excel.ActiveWorkbook);
            settings.prepend = "";
            txtPrepCode.Text = settings.prepend;
        }

    }
}
