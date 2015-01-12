using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

namespace FinalProject
{
    public partial class formReport : Form
    {
        public formReport()
        {
            InitializeComponent();
        }

        private void checkedListBoxReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkedListBoxReport.SelectedIndex == 0 && checkedListBoxReport.GetItemChecked(0))
            {
                for (int i = 0; i < checkedListBoxReport.Items.Count; i++)
                {
                    checkedListBoxReport.SetItemChecked(i, true);
                }
            }
            else if (checkedListBoxReport.SelectedIndex == 0 && !checkedListBoxReport.GetItemChecked(0))
            {
                for (int i = 0; i < checkedListBoxReport.Items.Count; i++)
                {
                    checkedListBoxReport.SetItemChecked(i, false);
                }
            }
            else 
            {
                checkedListBoxReport.SetItemChecked(0, false);
            }
        }

        private void btnGenerateReport_Click(object sender, EventArgs e)
        {
            string index = "";
            foreach (object itemChecked in checkedListBoxReport.CheckedIndices)
            {
                if (itemChecked.ToString() == "0")
                {
                    index += "0";
                    break;
                }
                else
                {
                    index += itemChecked.ToString();
                }
            }
            reportHandling.generateReport(index);
            MessageBox.Show(System.Windows.Forms.Application.StartupPath.ToString() + "\\Webconreport.docx","Word Report File Location",MessageBoxButtons.OK,MessageBoxIcon.Information);
            Process.Start("WebconReport.docx");
            this.Close();
        }

        private void formReport_Load(object sender, EventArgs e)
        {
            checkedListBoxReport.SetItemChecked(1, true);
        }
    }
}
