namespace mainframe_remake_two
{
    partial class Main_Form
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
            this.btnBrowseData = new System.Windows.Forms.Button();
            this.txtShowData = new System.Windows.Forms.TextBox();
            this.btnBrowseLookup = new System.Windows.Forms.Button();
            this.txtShowLookup = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.progReport = new System.Windows.Forms.ProgressBar();
            this.progWorker = new System.ComponentModel.BackgroundWorker();
            this.openFile = new System.Windows.Forms.OpenFileDialog();
            this.txtTitle = new System.Windows.Forms.TextBox();
            this.lblDivider = new System.Windows.Forms.Label();
            this.lblDivider2 = new System.Windows.Forms.Label();
            this.lblProgress = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnBrowseData
            // 
            this.btnBrowseData.Location = new System.Drawing.Point(194, 100);
            this.btnBrowseData.Name = "btnBrowseData";
            this.btnBrowseData.Size = new System.Drawing.Size(122, 40);
            this.btnBrowseData.TabIndex = 0;
            this.btnBrowseData.Text = "Open Data...";
            this.btnBrowseData.UseVisualStyleBackColor = true;
            this.btnBrowseData.Click += new System.EventHandler(this.BtnBrowseData_Click);
            // 
            // txtShowData
            // 
            this.txtShowData.Location = new System.Drawing.Point(322, 100);
            this.txtShowData.Multiline = true;
            this.txtShowData.Name = "txtShowData";
            this.txtShowData.ReadOnly = true;
            this.txtShowData.Size = new System.Drawing.Size(326, 40);
            this.txtShowData.TabIndex = 1;
            this.txtShowData.TabStop = false;
            // 
            // btnBrowseLookup
            // 
            this.btnBrowseLookup.Location = new System.Drawing.Point(194, 146);
            this.btnBrowseLookup.Name = "btnBrowseLookup";
            this.btnBrowseLookup.Size = new System.Drawing.Size(122, 40);
            this.btnBrowseLookup.TabIndex = 2;
            this.btnBrowseLookup.Text = "Open Lookup...";
            this.btnBrowseLookup.UseVisualStyleBackColor = true;
            this.btnBrowseLookup.Click += new System.EventHandler(this.BtnBrowseLookup_Click);
            // 
            // txtShowLookup
            // 
            this.txtShowLookup.Location = new System.Drawing.Point(322, 146);
            this.txtShowLookup.Multiline = true;
            this.txtShowLookup.Name = "txtShowLookup";
            this.txtShowLookup.ReadOnly = true;
            this.txtShowLookup.Size = new System.Drawing.Size(326, 40);
            this.txtShowLookup.TabIndex = 3;
            this.txtShowLookup.TabStop = false;
            // 
            // btnRun
            // 
            this.btnRun.Enabled = false;
            this.btnRun.Location = new System.Drawing.Point(194, 229);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(122, 40);
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "Run Report";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.BtnRun_Click);
            // 
            // progReport
            // 
            this.progReport.Location = new System.Drawing.Point(322, 244);
            this.progReport.Name = "progReport";
            this.progReport.Size = new System.Drawing.Size(326, 25);
            this.progReport.TabIndex = 5;
            // 
            // progWorker
            // 
            this.progWorker.WorkerReportsProgress = true;
            this.progWorker.WorkerSupportsCancellation = true;
            this.progWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.progWorker_DoWork);
            this.progWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.progWorker_ProgressChanged);
            this.progWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.progWorker_RunWorkerCompleted);
            // 
            // openFile
            // 
            this.openFile.Filter = "Excel Files (*.xls*)|*.xls*";
            this.openFile.Title = "Browse Excel Files";
            // 
            // txtTitle
            // 
            this.txtTitle.BackColor = System.Drawing.SystemColors.Control;
            this.txtTitle.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F);
            this.txtTitle.Location = new System.Drawing.Point(236, 27);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.ReadOnly = true;
            this.txtTitle.Size = new System.Drawing.Size(357, 31);
            this.txtTitle.TabIndex = 6;
            this.txtTitle.TabStop = false;
            this.txtTitle.Text = "Mainframe Reporting Script";
            this.txtTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblDivider
            // 
            this.lblDivider.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDivider.Location = new System.Drawing.Point(111, 73);
            this.lblDivider.Name = "lblDivider";
            this.lblDivider.Size = new System.Drawing.Size(600, 2);
            this.lblDivider.TabIndex = 7;
            // 
            // lblDivider2
            // 
            this.lblDivider2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDivider2.Location = new System.Drawing.Point(111, 205);
            this.lblDivider2.Name = "lblDivider2";
            this.lblDivider2.Size = new System.Drawing.Size(600, 2);
            this.lblDivider2.TabIndex = 8;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(322, 228);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(142, 13);
            this.lblProgress.TabIndex = 9;
            this.lblProgress.Text = "Awaiting Data and Lookup...";
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(806, 361);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.lblDivider2);
            this.Controls.Add(this.lblDivider);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.progReport);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.txtShowLookup);
            this.Controls.Add(this.btnBrowseLookup);
            this.Controls.Add(this.txtShowData);
            this.Controls.Add(this.btnBrowseData);
            this.MaximizeBox = false;
            this.Name = "Main_Form";
            this.ShowIcon = false;
            this.Text = "Mainframe Reporting Script";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowseData;
        private System.Windows.Forms.TextBox txtShowData;
        private System.Windows.Forms.Button btnBrowseLookup;
        private System.Windows.Forms.TextBox txtShowLookup;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.ProgressBar progReport;
        private System.ComponentModel.BackgroundWorker progWorker;
        private System.Windows.Forms.OpenFileDialog openFile;
        private System.Windows.Forms.TextBox txtTitle;
        private System.Windows.Forms.Label lblDivider;
        private System.Windows.Forms.Label lblDivider2;
        private System.Windows.Forms.Label lblProgress;
    }
}

