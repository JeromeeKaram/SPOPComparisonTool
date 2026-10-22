using System;
using System.IO;
using System.Windows.Forms;

namespace Cleaning_Records_Sheet.About
{
    public partial class UIHelp : Form
    {

        public const string pdf = "SPMComparison_CleaningRecordsHighlight.pdf";
        public const string pptx = "SPMComparison_CleaningRecordsHighlight.pptx";

        public UIHelp()
        {
            InitializeComponent();
        }

        private void UIHelp_Load(object sender, EventArgs e)
        {
            this.Text = this.Text = Utility.m_sToolName + " V" + Utility.m_sVersion;
            webBrowser1.Navigate(Path.Combine(Utility.m_sBinPath + "bin/", pdf));
        }

        private void pptToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filePath = Path.Combine(Utility.m_sBinPath, "bin/", pptx);

            if (File.Exists(filePath))
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true // <-- Important!
                };
                System.Diagnostics.Process.Start(psi);
            }
            else
            {
                MessageBox.Show("File not found:\n" + filePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //private void InitializeComponent()
        //{
        //    this.SuspendLayout();
        //    // 
        //    // UIHelp
        //    // 
        //    this.ClientSize = new System.Drawing.Size(284, 261);
        //    this.Name = "UIHelp";
        //    this.Load += new System.EventHandler(this.UIHelp_Load_1);
        //    this.ResumeLayout(false);

        //}

        private void UIHelp_Load_1(object sender, EventArgs e)
        {

        }
    }
}
