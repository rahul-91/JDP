namespace FinalProject
{
    partial class formReport
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
            this.checkedListBoxReport = new System.Windows.Forms.CheckedListBox();
            this.btnGenerateReport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // checkedListBoxReport
            // 
            this.checkedListBoxReport.BackColor = System.Drawing.SystemColors.Menu;
            this.checkedListBoxReport.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.checkedListBoxReport.CheckOnClick = true;
            this.checkedListBoxReport.ColumnWidth = 2;
            this.checkedListBoxReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkedListBoxReport.FormattingEnabled = true;
            this.checkedListBoxReport.Items.AddRange(new object[] {
            "Select All",
            "Sector-wise Distribution of Surveyed Units ",
            "Distribution of Surveyed Units By Category (All Sector)",
            "Ownership Status of Surveyed Units",
            "State-wise & Ownership-wise Distribution of JDP Units ",
            "Distribution of Annual Turnover of JDP Units during Last 3 Years",
            "Distribution Channels of JDP Units : 2012-13",
            "Power Consumption Scenario of Surveyed Units for Production of JDPs : 2012-13",
            "State Wise Distribution of Old and New JDP Units"});
            this.checkedListBoxReport.Location = new System.Drawing.Point(12, 35);
            this.checkedListBoxReport.Name = "checkedListBoxReport";
            this.checkedListBoxReport.Size = new System.Drawing.Size(499, 153);
            this.checkedListBoxReport.TabIndex = 1;
            this.checkedListBoxReport.UseTabStops = false;
            this.checkedListBoxReport.SelectedIndexChanged += new System.EventHandler(this.checkedListBoxReport_SelectedIndexChanged);
            // 
            // btnGenerateReport
            // 
            this.btnGenerateReport.Location = new System.Drawing.Point(122, 218);
            this.btnGenerateReport.Name = "btnGenerateReport";
            this.btnGenerateReport.Size = new System.Drawing.Size(235, 36);
            this.btnGenerateReport.TabIndex = 2;
            this.btnGenerateReport.Text = "Generate Word Report";
            this.btnGenerateReport.UseVisualStyleBackColor = true;
            this.btnGenerateReport.Click += new System.EventHandler(this.btnGenerateReport_Click);
            // 
            // formReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(518, 325);
            this.Controls.Add(this.btnGenerateReport);
            this.Controls.Add(this.checkedListBoxReport);
            this.Name = "formReport";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "formReport";
            this.Load += new System.EventHandler(this.formReport_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBoxReport;
        private System.Windows.Forms.Button btnGenerateReport;
    }
}