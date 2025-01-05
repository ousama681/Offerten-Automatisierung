namespace Offerten_Helper
{
    partial class AutomatedOffer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AutomatedOffer));
            TxtExcelFile = new TextBox();
            TxtPpptFile = new TextBox();
            CmdProcessPpt = new Button();
            CmdLoadExcelFile = new Button();
            CmdLoadPptFile = new Button();
            pictureBox2 = new PictureBox();
            pictureBox3 = new PictureBox();
            TxtStateDisplay = new TextBox();
            LstPptMissingNames = new ListBox();
            groupBox1 = new GroupBox();
            groupBox2 = new GroupBox();
            LblState = new Label();
            CmdCheckFiles = new Button();
            groupBox6 = new GroupBox();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).BeginInit();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox6.SuspendLayout();
            SuspendLayout();
            // 
            // TxtExcelFile
            // 
            TxtExcelFile.BackColor = SystemColors.Window;
            TxtExcelFile.ForeColor = SystemColors.WindowText;
            TxtExcelFile.Location = new Point(40, 22);
            TxtExcelFile.Name = "TxtExcelFile";
            TxtExcelFile.ReadOnly = true;
            TxtExcelFile.RightToLeft = RightToLeft.Yes;
            TxtExcelFile.Size = new Size(591, 23);
            TxtExcelFile.TabIndex = 8;
            // 
            // TxtPpptFile
            // 
            TxtPpptFile.BackColor = SystemColors.Window;
            TxtPpptFile.Location = new Point(40, 106);
            TxtPpptFile.Name = "TxtPpptFile";
            TxtPpptFile.ReadOnly = true;
            TxtPpptFile.RightToLeft = RightToLeft.Yes;
            TxtPpptFile.Size = new Size(591, 23);
            TxtPpptFile.TabIndex = 13;
            // 
            // CmdProcessPpt
            // 
            CmdProcessPpt.Enabled = false;
            CmdProcessPpt.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            CmdProcessPpt.Location = new Point(50, 330);
            CmdProcessPpt.Name = "CmdProcessPpt";
            CmdProcessPpt.Size = new Size(183, 42);
            CmdProcessPpt.TabIndex = 27;
            CmdProcessPpt.Text = "Create Offer";
            CmdProcessPpt.UseVisualStyleBackColor = true;
            CmdProcessPpt.Click += CmdProcessPpt_Click;
            // 
            // CmdLoadExcelFile
            // 
            CmdLoadExcelFile.Location = new Point(40, 51);
            CmdLoadExcelFile.Name = "CmdLoadExcelFile";
            CmdLoadExcelFile.Size = new Size(149, 30);
            CmdLoadExcelFile.TabIndex = 29;
            CmdLoadExcelFile.Text = "Load Excel File";
            CmdLoadExcelFile.UseVisualStyleBackColor = true;
            CmdLoadExcelFile.Click += CmdLoadExcelFile_Click;
            // 
            // CmdLoadPptFile
            // 
            CmdLoadPptFile.Location = new Point(40, 136);
            CmdLoadPptFile.Name = "CmdLoadPptFile";
            CmdLoadPptFile.Size = new Size(149, 30);
            CmdLoadPptFile.TabIndex = 30;
            CmdLoadPptFile.Text = "Load Powerpoint File";
            CmdLoadPptFile.UseVisualStyleBackColor = true;
            CmdLoadPptFile.Click += CmdLoadPptFile_Click;
            // 
            // pictureBox2
            // 
            pictureBox2.Image = Properties.Resources.PowerPoint_Logo_25_25;
            pictureBox2.Location = new Point(6, 106);
            pictureBox2.Name = "pictureBox2";
            pictureBox2.Size = new Size(27, 27);
            pictureBox2.TabIndex = 35;
            pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            pictureBox3.Image = Properties.Resources.Excel_Logo_25_25;
            pictureBox3.Location = new Point(6, 22);
            pictureBox3.Name = "pictureBox3";
            pictureBox3.Size = new Size(27, 25);
            pictureBox3.TabIndex = 36;
            pictureBox3.TabStop = false;
            // 
            // TxtStateDisplay
            // 
            TxtStateDisplay.BackColor = SystemColors.ControlLight;
            TxtStateDisplay.Enabled = false;
            TxtStateDisplay.Location = new Point(11, 22);
            TxtStateDisplay.Name = "TxtStateDisplay";
            TxtStateDisplay.ReadOnly = true;
            TxtStateDisplay.Size = new Size(23, 23);
            TxtStateDisplay.TabIndex = 37;
            // 
            // LstPptMissingNames
            // 
            LstPptMissingNames.FormattingEnabled = true;
            LstPptMissingNames.ItemHeight = 15;
            LstPptMissingNames.Location = new Point(6, 22);
            LstPptMissingNames.Name = "LstPptMissingNames";
            LstPptMissingNames.Size = new Size(249, 139);
            LstPptMissingNames.TabIndex = 39;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(CmdLoadExcelFile);
            groupBox1.Controls.Add(TxtExcelFile);
            groupBox1.Controls.Add(TxtPpptFile);
            groupBox1.Controls.Add(CmdLoadPptFile);
            groupBox1.Controls.Add(pictureBox2);
            groupBox1.Controls.Add(pictureBox3);
            groupBox1.Location = new Point(10, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(640, 185);
            groupBox1.TabIndex = 43;
            groupBox1.TabStop = false;
            groupBox1.Text = "File Selection";
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(LblState);
            groupBox2.Controls.Add(CmdCheckFiles);
            groupBox2.Controls.Add(TxtStateDisplay);
            groupBox2.FlatStyle = FlatStyle.Flat;
            groupBox2.Location = new Point(10, 211);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(369, 94);
            groupBox2.TabIndex = 44;
            groupBox2.TabStop = false;
            groupBox2.Text = "Status";
            // 
            // LblState
            // 
            LblState.AutoSize = true;
            LblState.Location = new Point(40, 25);
            LblState.Name = "LblState";
            LblState.Size = new Size(202, 15);
            LblState.TabIndex = 38;
            LblState.Text = "Choose both files to enable checking";
            // 
            // CmdCheckFiles
            // 
            CmdCheckFiles.Enabled = false;
            CmdCheckFiles.Location = new Point(40, 51);
            CmdCheckFiles.Name = "CmdCheckFiles";
            CmdCheckFiles.Size = new Size(143, 30);
            CmdCheckFiles.TabIndex = 37;
            CmdCheckFiles.Text = "Check Files";
            CmdCheckFiles.UseVisualStyleBackColor = true;
            CmdCheckFiles.Click += CmdCheckFiles_Click;
            // 
            // groupBox6
            // 
            groupBox6.Controls.Add(LstPptMissingNames);
            groupBox6.Location = new Point(390, 211);
            groupBox6.Name = "groupBox6";
            groupBox6.Size = new Size(260, 169);
            groupBox6.TabIndex = 49;
            groupBox6.TabStop = false;
            groupBox6.Text = "Missing Names in PowerPoint";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(659, 388);
            Controls.Add(groupBox6);
            Controls.Add(groupBox2);
            Controls.Add(CmdProcessPpt);
            Controls.Add(groupBox1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "Form1";
            Text = "Automated Offer";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox6.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion
        private TextBox TxtExcelFile;
        private TextBox TxtPpptFile;
        private Button CmdProcessPpt;
        private Button CmdLoadExcelFile;
        private Button CmdLoadPptFile;
        private PictureBox pictureBox2;
        private PictureBox pictureBox3;
        private TextBox TxtStateDisplay;
        private ListBox LstPptMissingNames;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private Button CmdCheckFiles;
        private GroupBox groupBox6;
        private Label LblState;
    }
}
