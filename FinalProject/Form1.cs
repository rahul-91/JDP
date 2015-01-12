using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace FinalProject
{
    public partial class formImport : Form
    {
        public formImport()
        {
            InitializeComponent();
        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            btnImport.Enabled = false;
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Import Excel Input File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "Microsoft Excel files (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                timer1.Start();
                progressBarImport.Visible = true;
                excelHandling.readExcelfile(fdlg.FileName.Replace(@"\", @"\\").ToString());
                formReport report = new formReport();
                report.Show();
            }
            progressBarImport.Value = 99;
            btnImport.Enabled = true;
            progressBarImport.Visible = false;
            progressBarImport.Value = 0;
        }

        private void formImport_Load(object sender, EventArgs e)
        {
            progressBarImport.Visible = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBarImport.Increment(1);
        }
    }
}
