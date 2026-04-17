namespace FormTranslator
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            folderBrowserDialog1 = new FolderBrowserDialog();
            lblFolder = new Label();
            txtFolderPath = new TextBox();
            btnBrowse = new Button();
            lblLanguage = new Label();
            cboLanguage = new ComboBox();
            chkExportCsv = new CheckBox();
            chkConvertCodeBehind = new CheckBox();
            btnTranslate = new Button();
            progressBar1 = new ProgressBar();
            lblProgress = new Label();
            txtLog = new TextBox();
            SuspendLayout();
            // 
            // lblFolder
            // 
            lblFolder.AutoSize = true;
            lblFolder.Location = new Point(12, 15);
            lblFolder.Name = "lblFolder";
            lblFolder.Size = new Size(234, 25);
            lblFolder.TabIndex = 0;
            lblFolder.Text = "Folder of forms to translate:";
            // 
            // txtFolderPath
            // 
            txtFolderPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtFolderPath.Location = new Point(252, 12);
            txtFolderPath.Name = "txtFolderPath";
            txtFolderPath.Size = new Size(556, 31);
            txtFolderPath.TabIndex = 1;
            // 
            // btnBrowse
            // 
            btnBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnBrowse.Location = new Point(814, 11);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(112, 34);
            btnBrowse.TabIndex = 2;
            btnBrowse.Text = "Browse";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += BtnBrowse_Click;
            // 
            // lblLanguage
            // 
            lblLanguage.AutoSize = true;
            lblLanguage.Location = new Point(12, 61);
            lblLanguage.Name = "lblLanguage";
            lblLanguage.Size = new Size(93, 25);
            lblLanguage.TabIndex = 3;
            lblLanguage.Text = "Language:";
            // 
            // cboLanguage
            // 
            cboLanguage.DropDownStyle = ComboBoxStyle.DropDownList;
            cboLanguage.FormattingEnabled = true;
            cboLanguage.Location = new Point(117, 57);
            cboLanguage.Name = "cboLanguage";
            cboLanguage.Size = new Size(268, 33);
            cboLanguage.TabIndex = 4;
            cboLanguage.SelectedIndexChanged += CboLanguage_SelectedIndexChanged;
            // 
            // chkExportCsv
            // 
            chkExportCsv.AutoSize = true;
            chkExportCsv.Checked = true;
            chkExportCsv.CheckState = CheckState.Checked;
            chkExportCsv.Location = new Point(391, 60);
            chkExportCsv.Name = "chkExportCsv";
            chkExportCsv.Size = new Size(183, 29);
            chkExportCsv.TabIndex = 5;
            chkExportCsv.Text = "Export folder CSVs";
            chkExportCsv.UseVisualStyleBackColor = true;
            chkExportCsv.CheckedChanged += ChkExportCsv_CheckedChanged;
            // 
            // chkConvertCodeBehind
            // 
            chkConvertCodeBehind.AutoSize = true;
            chkConvertCodeBehind.Location = new Point(580, 60);
            chkConvertCodeBehind.Name = "chkConvertCodeBehind";
            chkConvertCodeBehind.Size = new Size(193, 29);
            chkConvertCodeBehind.TabIndex = 6;
            chkConvertCodeBehind.Text = "Conver code behind";
            chkConvertCodeBehind.UseVisualStyleBackColor = true;
            chkConvertCodeBehind.CheckedChanged += ChkConvertCodeBehind_CheckedChanged;
            // 
            // btnTranslate
            // 
            btnTranslate.Location = new Point(779, 56);
            btnTranslate.Name = "btnTranslate";
            btnTranslate.Size = new Size(147, 34);
            btnTranslate.TabIndex = 7;
            btnTranslate.Text = "Translate to es-MX";
            btnTranslate.UseVisualStyleBackColor = true;
            btnTranslate.Click += BtnTranslate_Click;
            // 
            // progressBar1
            // 
            progressBar1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            progressBar1.Location = new Point(12, 96);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(846, 30);
            progressBar1.TabIndex = 8;
            // 
            // lblProgress
            // 
            lblProgress.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            lblProgress.Location = new Point(871, 96);
            lblProgress.Name = "lblProgress";
            lblProgress.Size = new Size(55, 30);
            lblProgress.TabIndex = 9;
            lblProgress.Text = "0%";
            lblProgress.TextAlign = ContentAlignment.MiddleRight;
            // 
            // txtLog
            // 
            txtLog.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtLog.Location = new Point(12, 132);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new Size(914, 271);
            txtLog.TabIndex = 10;
            txtLog.WordWrap = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(938, 415);
            Controls.Add(txtLog);
            Controls.Add(lblProgress);
            Controls.Add(progressBar1);
            Controls.Add(btnTranslate);
            Controls.Add(chkConvertCodeBehind);
            Controls.Add(chkExportCsv);
            Controls.Add(cboLanguage);
            Controls.Add(lblLanguage);
            Controls.Add(btnBrowse);
            Controls.Add(txtFolderPath);
            Controls.Add(lblFolder);
            MinimumSize = new Size(900, 450);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "WinForms - Form UI Translator V 1.0";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private FolderBrowserDialog folderBrowserDialog1;
        private Label lblFolder;
        private TextBox txtFolderPath;
        private Button btnBrowse;
        private Label lblLanguage;
        private ComboBox cboLanguage;
        private CheckBox chkExportCsv;
        private CheckBox chkConvertCodeBehind;
        private Button btnTranslate;
        private ProgressBar progressBar1;
        private Label lblProgress;
        private TextBox txtLog;
    }
}
