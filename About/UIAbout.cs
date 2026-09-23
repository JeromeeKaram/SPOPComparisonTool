using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SVC.About
{
    public partial class UIAbout : Form
    {
        public UIAbout()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void UIAbout_Load(object sender, EventArgs e)
        {
            try
            {
                lbVersion.Text = "Version : " + Utility.m_sVersion;
                lbDate.Text = "Last Release : " + Utility.m_sDate;
                this.Text = this.Text = Utility.m_sToolName + " V" + Utility.m_sVersion;
                //if (type == 0)
                //{
                //    label3.Text = "The main purpose of the tool is to automate the process of comparing New revision of SB tasks in engine manual with Old revision of engine manual";
                //}

                List<string> lst = Utility.ReadFile(Utility.m_sBinPath + "bin/" + "contact.ini");
                List<string> lst1 = Utility.SplitString(lst[0], ";");
                List<string> lst2 = Utility.SplitString(lst[1], ";");
                List<string> lst3 = Utility.SplitString(lst[2], ";");
                lblHIContact.Text = lst1[0];
                tbIHIMail.Text = lst1[1];
                lbOnsiteContact.Text = lst2[0];
                tbOnsiteMail.Text = lst2[1];
                lbOffshoreContact.Text = lst3[0];
                tbOffshoreMail.Text = lst3[1];
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
    }
}
