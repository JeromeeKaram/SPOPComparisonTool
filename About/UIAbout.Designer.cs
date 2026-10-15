using System.Drawing;
using System.Windows.Forms;

namespace Cleaning_Records_Sheet.About
{
    partial class UIAbout
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            panel1 = new Panel();
            label3 = new Label();
            grbContactDet = new GroupBox();
            tbOffshoreMail = new TextBox();
            tbOnsiteMail = new TextBox();
            tbIHIMail = new TextBox();
            lbOffshoreContact = new Label();
            lbOnsiteContact = new Label();
            lblHIContact = new Label();
            label2 = new Label();
            label1 = new Label();
            lbContact = new Label();
            BtnOk = new Button();
            lbVersion = new Label();
            lbDate = new Label();
            panel1.SuspendLayout();
            grbContactDet.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.BackColor = SystemColors.GradientActiveCaption;
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel1.Controls.Add(label3);
            panel1.Font = new Font("Verdana", 9F, FontStyle.Italic);
            panel1.Location = new Point(5, 36);
            panel1.Margin = new Padding(4, 3, 4, 3);
            panel1.Name = "panel1";
            panel1.Size = new Size(690, 80);
            panel1.TabIndex = 0;
            // 
            // label3
            // 
            label3.BackColor = SystemColors.GradientActiveCaption;
            label3.Font = new Font("Verdana", 9F);
            label3.ForeColor = Color.FromArgb(25, 70, 130);
            label3.Location = new Point(6, 15);
            label3.Margin = new Padding(4, 0, 4, 0);
            label3.Name = "label3";
            label3.Size = new Size(680, 68);
            label3.TabIndex = 9;
            label3.Text = "The primary purpose of this tool is to compare previous and latest Engine Maintenance Manual (SPOP) revisions against Cleaning Record requirements and automatically highlight any procedure changes, missing requirements, or discrepancies in the Cleaning Record Excel file.";
            // 
            // grbContactDet
            // 
            grbContactDet.BackColor = SystemColors.GradientActiveCaption;
            grbContactDet.Controls.Add(tbOffshoreMail);
            grbContactDet.Controls.Add(tbOnsiteMail);
            grbContactDet.Controls.Add(tbIHIMail);
            grbContactDet.Controls.Add(lbOffshoreContact);
            grbContactDet.Controls.Add(lbOnsiteContact);
            grbContactDet.Controls.Add(lblHIContact);
            grbContactDet.Controls.Add(label2);
            grbContactDet.Controls.Add(label1);
            grbContactDet.Controls.Add(lbContact);
            grbContactDet.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 254);
            grbContactDet.ForeColor = Color.FromArgb(25, 70, 130);
            grbContactDet.Location = new Point(5, 123);
            grbContactDet.Margin = new Padding(4, 3, 4, 3);
            grbContactDet.Name = "grbContactDet";
            grbContactDet.Padding = new Padding(4, 3, 4, 3);
            grbContactDet.Size = new Size(691, 140);
            grbContactDet.TabIndex = 1;
            grbContactDet.TabStop = false;
            grbContactDet.Text = "Contact details";
            // 
            // tbOffshoreMail
            // 
            tbOffshoreMail.BackColor = SystemColors.GradientActiveCaption;
            tbOffshoreMail.BorderStyle = BorderStyle.None;
            tbOffshoreMail.Font = new Font("Bell MT", 8.25F, FontStyle.Underline);
            tbOffshoreMail.ForeColor = Color.FromArgb(25, 70, 130);
            tbOffshoreMail.Location = new Point(387, 107);
            tbOffshoreMail.Margin = new Padding(4, 3, 4, 3);
            tbOffshoreMail.Name = "tbOffshoreMail";
            tbOffshoreMail.Size = new Size(216, 13);
            tbOffshoreMail.TabIndex = 11;
            tbOffshoreMail.Text = "Veerendra.Kotari@cyient.com";
            // 
            // tbOnsiteMail
            // 
            tbOnsiteMail.BackColor = SystemColors.GradientActiveCaption;
            tbOnsiteMail.BorderStyle = BorderStyle.None;
            tbOnsiteMail.Font = new Font("Bell MT", 8.25F, FontStyle.Underline);
            tbOnsiteMail.ForeColor = Color.FromArgb(25, 70, 130);
            tbOnsiteMail.Location = new Point(387, 70);
            tbOnsiteMail.Margin = new Padding(4, 3, 4, 3);
            tbOnsiteMail.Name = "tbOnsiteMail";
            tbOnsiteMail.Size = new Size(216, 13);
            tbOnsiteMail.TabIndex = 10;
            tbOnsiteMail.Text = "mahendra5603@ihi-g.com";
            // 
            // tbIHIMail
            // 
            tbIHIMail.BackColor = SystemColors.GradientActiveCaption;
            tbIHIMail.BorderStyle = BorderStyle.None;
            tbIHIMail.Font = new Font("Bell MT", 8.25F, FontStyle.Underline);
            tbIHIMail.ForeColor = Color.FromArgb(25, 70, 130);
            tbIHIMail.Location = new Point(387, 32);
            tbIHIMail.Margin = new Padding(4, 3, 4, 3);
            tbIHIMail.Name = "tbIHIMail";
            tbIHIMail.Size = new Size(216, 13);
            tbIHIMail.TabIndex = 9;
            tbIHIMail.Text = "hayasaka1386@ihi-g.com";
            // 
            // lbOffshoreContact
            // 
            lbOffshoreContact.AutoSize = true;
            lbOffshoreContact.Font = new Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lbOffshoreContact.Location = new Point(178, 106);
            lbOffshoreContact.Margin = new Padding(4, 0, 4, 0);
            lbOffshoreContact.Name = "lbOffshoreContact";
            lbOffshoreContact.Size = new Size(119, 14);
            lbOffshoreContact.TabIndex = 5;
            lbOffshoreContact.Text = "Veerendra Kotari";
            // 
            // lbOnsiteContact
            // 
            lbOnsiteContact.AutoSize = true;
            lbOnsiteContact.Font = new Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lbOnsiteContact.Location = new Point(178, 69);
            lbOnsiteContact.Margin = new Padding(4, 0, 4, 0);
            lbOnsiteContact.Name = "lbOnsiteContact";
            lbOnsiteContact.Size = new Size(112, 14);
            lbOnsiteContact.TabIndex = 4;
            lbOnsiteContact.Text = "Guraka Mahendra";
            // 
            // lblHIContact
            // 
            lblHIContact.AutoSize = true;
            lblHIContact.Font = new Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblHIContact.Location = new Point(178, 32);
            lblHIContact.Margin = new Padding(4, 0, 4, 0);
            lblHIContact.Name = "lblHIContact";
            lblHIContact.Size = new Size(91, 14);
            lblHIContact.TabIndex = 3;
            lblHIContact.Text = "Hayasaka san";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Courier New", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label2.Location = new Point(19, 106);
            label2.Margin = new Padding(4, 0, 4, 0);
            label2.Name = "label2";
            label2.Size = new Size(112, 14);
            label2.TabIndex = 2;
            label2.Text = "CYIENT Offshore";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Courier New", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label1.Location = new Point(19, 69);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new Size(98, 14);
            label1.TabIndex = 1;
            label1.Text = "CYIENT Onsite";
            // 
            // lbContact
            // 
            lbContact.AutoSize = true;
            lbContact.Font = new Font("Courier New", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lbContact.Location = new Point(19, 32);
            lbContact.Margin = new Padding(4, 0, 4, 0);
            lbContact.Name = "lbContact";
            lbContact.Size = new Size(28, 14);
            lbContact.TabIndex = 0;
            lbContact.Text = "IHI";
            // 
            // BtnOk
            // 
            BtnOk.Location = new Point(588, 273);
            BtnOk.Margin = new Padding(4, 3, 4, 3);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new Size(107, 37);
            BtnOk.TabIndex = 0;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // lbVersion
            // 
            lbVersion.AutoSize = true;
            lbVersion.BackColor = SystemColors.GradientActiveCaption;
            lbVersion.Font = new Font("Verdana", 9F, FontStyle.Italic, GraphicsUnit.Point, 0);
            lbVersion.ForeColor = Color.FromArgb(25, 70, 130);
            lbVersion.Location = new Point(9, 10);
            lbVersion.Margin = new Padding(4, 0, 4, 0);
            lbVersion.Name = "lbVersion";
            lbVersion.Size = new Size(52, 14);
            lbVersion.TabIndex = 57;
            lbVersion.Text = "Version";
            // 
            // lbDate
            // 
            lbDate.AutoSize = true;
            lbDate.BackColor = SystemColors.GradientActiveCaption;
            lbDate.Font = new Font("Verdana", 9F, FontStyle.Italic, GraphicsUnit.Point, 0);
            lbDate.ForeColor = Color.FromArgb(25, 70, 130);
            lbDate.Location = new Point(142, 10);
            lbDate.Margin = new Padding(4, 0, 4, 0);
            lbDate.Name = "lbDate";
            lbDate.Size = new Size(35, 14);
            lbDate.TabIndex = 58;
            lbDate.Text = "Date";
            // 
            // UIAbout
            // 
            AcceptButton = BtnOk;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.GradientActiveCaption;
            ClientSize = new Size(702, 321);
            Controls.Add(lbDate);
            Controls.Add(lbVersion);
            Controls.Add(BtnOk);
            Controls.Add(grbContactDet);
            Controls.Add(panel1);
            FormBorderStyle = FormBorderStyle.Fixed3D;
            KeyPreview = true;
            Margin = new Padding(4, 3, 4, 3);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "UIAbout";
            ShowInTaskbar = false;
            Text = "SVC Comparison Tool V1.0";
            Load += UIAbout_Load;
            panel1.ResumeLayout(false);
            grbContactDet.ResumeLayout(false);
            grbContactDet.PerformLayout();
            ResumeLayout(false);
            PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox grbContactDet;
        private System.Windows.Forms.Label lbContact;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lbVersion;
        private System.Windows.Forms.Label lbDate;
        private System.Windows.Forms.Label lbOffshoreContact;
        private System.Windows.Forms.Label lbOnsiteContact;
        private System.Windows.Forms.Label lblHIContact;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbIHIMail;
        private System.Windows.Forms.TextBox tbOffshoreMail;
        private System.Windows.Forms.TextBox tbOnsiteMail;
    }
}