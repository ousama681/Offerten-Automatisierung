namespace Offerten_Helper
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
            components = new System.ComponentModel.Container();
            LblExcelFile = new Label();
            LblPptFile = new Label();
            LblSaveLocation = new Label();
            TxtExcelFile = new TextBox();
            TxtPpptFile = new TextBox();
            TxtSaveLocation = new TextBox();
            LblMappingfileName = new Label();
            CmdProcessPpt = new Button();
            CmdLoadExcelFile = new Button();
            CmdLoadPptFile = new Button();
            CmdSaveLocation = new Button();
            errorProvider1 = new ErrorProvider(components);
            pictureBox1 = new PictureBox();
            pictureBox2 = new PictureBox();
            pictureBox3 = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)errorProvider1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).BeginInit();
            SuspendLayout();
            // 
            // LblExcelFile
            // 
            LblExcelFile.AutoSize = true;
            LblExcelFile.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            LblExcelFile.Location = new Point(228, 315);
            LblExcelFile.Name = "LblExcelFile";
            LblExcelFile.Size = new Size(71, 21);
            LblExcelFile.TabIndex = 0;
            LblExcelFile.Text = "ExcelFile:";
            // 
            // LblPptFile
            // 
            LblPptFile.AutoSize = true;
            LblPptFile.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            LblPptFile.Location = new Point(490, 333);
            LblPptFile.Name = "LblPptFile";
            LblPptFile.Size = new Size(121, 21);
            LblPptFile.TabIndex = 1;
            LblPptFile.Text = "PowerPoint File:";
            // 
            // LblSaveLocation
            // 
            LblSaveLocation.AutoSize = true;
            LblSaveLocation.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            LblSaveLocation.Location = new Point(314, 333);
            LblSaveLocation.Name = "LblSaveLocation";
            LblSaveLocation.Size = new Size(109, 21);
            LblSaveLocation.TabIndex = 4;
            LblSaveLocation.Text = "Save Location:";
            // 
            // TxtExcelFile
            // 
            TxtExcelFile.Enabled = false;
            TxtExcelFile.Location = new Point(40, 12);
            TxtExcelFile.Name = "TxtExcelFile";
            TxtExcelFile.ReadOnly = true;
            TxtExcelFile.Size = new Size(591, 23);
            TxtExcelFile.TabIndex = 8;
            // 
            // TxtPpptFile
            // 
            TxtPpptFile.Enabled = false;
            TxtPpptFile.Location = new Point(38, 86);
            TxtPpptFile.Name = "TxtPpptFile";
            TxtPpptFile.ReadOnly = true;
            TxtPpptFile.Size = new Size(591, 23);
            TxtPpptFile.TabIndex = 13;
            // 
            // TxtSaveLocation
            // 
            TxtSaveLocation.Enabled = false;
            TxtSaveLocation.Location = new Point(40, 384);
            TxtSaveLocation.Name = "TxtSaveLocation";
            TxtSaveLocation.ReadOnly = true;
            TxtSaveLocation.Size = new Size(591, 23);
            TxtSaveLocation.TabIndex = 13;
            // 
            // LblMappingfileName
            // 
            LblMappingfileName.AutoSize = true;
            LblMappingfileName.Location = new Point(111, 164);
            LblMappingfileName.Name = "LblMappingfileName";
            LblMappingfileName.Size = new Size(0, 15);
            LblMappingfileName.TabIndex = 25;
            // 
            // CmdProcessPpt
            // 
            CmdProcessPpt.Enabled = false;
            CmdProcessPpt.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            CmdProcessPpt.Location = new Point(487, 115);
            CmdProcessPpt.Name = "CmdProcessPpt";
            CmdProcessPpt.Size = new Size(142, 38);
            CmdProcessPpt.TabIndex = 27;
            CmdProcessPpt.Text = "Create Offer";
            CmdProcessPpt.UseVisualStyleBackColor = true;
            CmdProcessPpt.Click += CmdProcessPpt_Click;
            // 
            // CmdLoadExcelFile
            // 
            CmdLoadExcelFile.Location = new Point(40, 41);
            CmdLoadExcelFile.Name = "CmdLoadExcelFile";
            CmdLoadExcelFile.Size = new Size(149, 23);
            CmdLoadExcelFile.TabIndex = 29;
            CmdLoadExcelFile.Text = "Load Excel File";
            CmdLoadExcelFile.UseVisualStyleBackColor = true;
            CmdLoadExcelFile.Click += CmdLoadExcelFile_Click;
            // 
            // CmdLoadPptFile
            // 
            CmdLoadPptFile.Location = new Point(38, 115);
            CmdLoadPptFile.Name = "CmdLoadPptFile";
            CmdLoadPptFile.Size = new Size(149, 23);
            CmdLoadPptFile.TabIndex = 30;
            CmdLoadPptFile.Text = "Load Powerpoint File";
            CmdLoadPptFile.UseVisualStyleBackColor = true;
            CmdLoadPptFile.Click += CmdLoadPptFile_Click;
            // 
            // CmdSaveLocation
            // 
            CmdSaveLocation.Font = new Font("Ebrima", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            CmdSaveLocation.Location = new Point(38, 413);
            CmdSaveLocation.Name = "CmdSaveLocation";
            CmdSaveLocation.Size = new Size(149, 24);
            CmdSaveLocation.TabIndex = 33;
            CmdSaveLocation.Text = "Choose Save Location";
            CmdSaveLocation.UseVisualStyleBackColor = true;
            CmdSaveLocation.Click += CmdSaveLocation_Click;
            // 
            // errorProvider1
            // 
            errorProvider1.ContainerControl = this;
            // 
            // pictureBox1
            // 
            errorProvider1.SetIconAlignment(pictureBox1, ErrorIconAlignment.TopLeft);
            pictureBox1.Image = Properties.Resources.FloppyDisk_Logo_25_25;
            pictureBox1.Location = new Point(11, 380);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(26, 26);
            pictureBox1.TabIndex = 34;
            pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            pictureBox2.Image = Properties.Resources.PowerPoint_Logo_25_25;
            pictureBox2.Location = new Point(10, 85);
            pictureBox2.Name = "pictureBox2";
            pictureBox2.Size = new Size(27, 27);
            pictureBox2.TabIndex = 35;
            pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            pictureBox3.Image = Properties.Resources.Excel_Logo_25_25;
            pictureBox3.Location = new Point(11, 12);
            pictureBox3.Name = "pictureBox3";
            pictureBox3.Size = new Size(27, 25);
            pictureBox3.TabIndex = 36;
            pictureBox3.TabStop = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(640, 165);
            Controls.Add(pictureBox3);
            Controls.Add(pictureBox2);
            Controls.Add(pictureBox1);
            Controls.Add(CmdSaveLocation);
            Controls.Add(CmdLoadPptFile);
            Controls.Add(CmdLoadExcelFile);
            Controls.Add(CmdProcessPpt);
            Controls.Add(LblMappingfileName);
            Controls.Add(TxtSaveLocation);
            Controls.Add(TxtPpptFile);
            Controls.Add(TxtExcelFile);
            Controls.Add(LblSaveLocation);
            Controls.Add(LblPptFile);
            Controls.Add(LblExcelFile);
            MaximizeBox = false;
            Name = "Form1";
            Text = "Automated Offer";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)errorProvider1).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label LblExcelFile;
        private Label LblPptFile;
        private Label LblSaveLocation;
        private TextBox TxtExcelFile;
        private TextBox TxtPpptFile;
        private TextBox TxtSaveLocation;
        private Label LblMappingfileName;
        private Button CmdProcessPpt;
        private Button CmdLoadExcelFile;
        private Button CmdLoadPptFile;
        private Button CmdSaveLocation;
        private ErrorProvider errorProvider1;
        private PictureBox pictureBox1;
        private PictureBox pictureBox2;
        private PictureBox pictureBox3;
    }
}
