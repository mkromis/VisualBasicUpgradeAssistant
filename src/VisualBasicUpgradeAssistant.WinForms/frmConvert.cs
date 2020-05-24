using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Data;
using VisualBasicUpgradeAssistant.WinForms;
using VisualBasicUpgradeAssistant.Core.Model;

namespace VB2C
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class FrmConvert : System.Windows.Forms.Form
    {
        private System.Windows.Forms.TextBox txtVB6;
        private System.Windows.Forms.Button cmdLoad;
        private System.Windows.Forms.Button cmdConvert;
        private System.Windows.Forms.Button cmdExit;
        private System.Windows.Forms.TextBox txtCSharp;
        private System.Windows.Forms.Label label;
        private System.Windows.Forms.TextBox txtOutPath;
        private System.Windows.Forms.Label label1;
        private System.ComponentModel.IContainer components = null;

        public FrmConvert()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmConvert));
            this.cmdExit = new System.Windows.Forms.Button();
            this.label = new System.Windows.Forms.Label();
            this.txtCSharp = new System.Windows.Forms.TextBox();
            this.cmdConvert = new System.Windows.Forms.Button();
            this.cmdLoad = new System.Windows.Forms.Button();
            this.txtVB6 = new System.Windows.Forms.TextBox();
            this.txtOutPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmdExit
            // 
            this.cmdExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdExit.BackColor = System.Drawing.SystemColors.ControlLight;
            this.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdExit.Location = new System.Drawing.Point(1471, 1275);
            this.cmdExit.Name = "cmdExit";
            this.cmdExit.Size = new System.Drawing.Size(192, 98);
            this.cmdExit.TabIndex = 0;
            this.cmdExit.Text = "Exit";
            this.cmdExit.UseVisualStyleBackColor = false;
            this.cmdExit.Click += new System.EventHandler(this.cmdExit_Click);
            // 
            // label
            // 
            this.label.Location = new System.Drawing.Point(40, 16);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(128, 24);
            this.label.TabIndex = 0;
            this.label.Text = "label";
            // 
            // txtCSharp
            // 
            this.txtCSharp.AcceptsReturn = true;
            this.txtCSharp.AcceptsTab = true;
            this.txtCSharp.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtCSharp.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtCSharp.Location = new System.Drawing.Point(15, 648);
            this.txtCSharp.MaxLength = 327670;
            this.txtCSharp.Multiline = true;
            this.txtCSharp.Name = "txtCSharp";
            this.txtCSharp.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCSharp.Size = new System.Drawing.Size(1651, 610);
            this.txtCSharp.TabIndex = 3;
            this.txtCSharp.WordWrap = false;
            // 
            // cmdConvert
            // 
            this.cmdConvert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdConvert.Location = new System.Drawing.Point(240, 1275);
            this.cmdConvert.Name = "cmdConvert";
            this.cmdConvert.Size = new System.Drawing.Size(192, 98);
            this.cmdConvert.TabIndex = 6;
            this.cmdConvert.Text = "Convert";
            this.cmdConvert.Click += new System.EventHandler(this.cmdConvert_Click);
            // 
            // cmdLoad
            // 
            this.cmdLoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdLoad.Location = new System.Drawing.Point(19, 1275);
            this.cmdLoad.Name = "cmdLoad";
            this.cmdLoad.Size = new System.Drawing.Size(192, 98);
            this.cmdLoad.TabIndex = 4;
            this.cmdLoad.Text = "Load";
            this.cmdLoad.Click += new System.EventHandler(this.cmdLoad_Click);
            // 
            // txtVB6
            // 
            this.txtVB6.AcceptsReturn = true;
            this.txtVB6.AcceptsTab = true;
            this.txtVB6.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtVB6.Font = new System.Drawing.Font("Courier New", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtVB6.Location = new System.Drawing.Point(15, 18);
            this.txtVB6.MaxLength = 327670;
            this.txtVB6.Multiline = true;
            this.txtVB6.Name = "txtVB6";
            this.txtVB6.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtVB6.Size = new System.Drawing.Size(1651, 610);
            this.txtVB6.TabIndex = 2;
            this.txtVB6.WordWrap = false;
            // 
            // txtOutPath
            // 
            this.txtOutPath.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.txtOutPath.Location = new System.Drawing.Point(608, 1300);
            this.txtOutPath.Name = "txtOutPath";
            this.txtOutPath.Size = new System.Drawing.Size(528, 39);
            this.txtOutPath.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(452, 1300);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 32);
            this.label1.TabIndex = 8;
            this.label1.Text = "Out path";
            // 
            // frmConvert
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(12, 32);
            this.CancelButton = this.cmdExit;
            this.ClientSize = new System.Drawing.Size(1683, 1386);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtOutPath);
            this.Controls.Add(this.cmdConvert);
            this.Controls.Add(this.cmdLoad);
            this.Controls.Add(this.txtCSharp);
            this.Controls.Add(this.txtVB6);
            this.Controls.Add(this.cmdExit);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(1709, 1457);
            this.Name = "frmConvert";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Convert VB6 to C#";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmConvert_Closing);
            this.Load += new System.EventHandler(this.frmConvert_Load);
            this.Resize += new System.EventHandler(this.frmConvert_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private string mFileName = null;
        private string mOutPath = null;

        private void cmdExit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void cmdConvert_Click(object sender, System.EventArgs e)
        {

        }

        private void frmConvert_Load(object sender, System.EventArgs e)
        {
            txtOutPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private void frmConvert_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Don't write out to program directory
            //Program.Config.WriteString(Program.CONFIG_SETTING, Program.CONFIG_OUT_PATH, txtOutPath.Text);
        }

        //    private string FileSave()
        //    {
        //      string sFilter = "C# Files (*.cs)|*.cs" ;	
        //      string sResult = null;
        //
        //      SaveFileDialog oDialog = new SaveFileDialog();		
        //      oDialog.Filter = sFilter;
        //      if(oDialog.ShowDialog() != DialogResult.Cancel)	
        //      {		
        //        sResult = oDialog.FileName;
        //      }	
        //      return sResult;
        //    }


    }
}
