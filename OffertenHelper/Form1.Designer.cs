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
            ImgExcelLogo = new PictureBox();
            ImgPowerPointLogo = new PictureBox();
            CmdSaveLocation = new Button();
            ImgSaveDiskLogo = new PictureBox();
            errorProvider1 = new ErrorProvider(components);
            ((System.ComponentModel.ISupportInitialize)ImgExcelLogo).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ImgPowerPointLogo).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ImgSaveDiskLogo).BeginInit();
            ((System.ComponentModel.ISupportInitialize)errorProvider1).BeginInit();
            SuspendLayout();
            // 
            // LblExcelFile
            // 
            LblExcelFile.AutoSize = true;
            LblExcelFile.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            LblExcelFile.Location = new Point(63, 11);
            LblExcelFile.Name = "LblExcelFile";
            LblExcelFile.Size = new Size(71, 21);
            LblExcelFile.TabIndex = 0;
            LblExcelFile.Text = "ExcelFile:";
            // 
            // LblPptFile
            // 
            LblPptFile.AutoSize = true;
            LblPptFile.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            LblPptFile.Location = new Point(63, 77);
            LblPptFile.Name = "LblPptFile";
            LblPptFile.Size = new Size(121, 21);
            LblPptFile.TabIndex = 1;
            LblPptFile.Text = "PowerPoint File:";
            // 
            // LblSaveLocation
            // 
            LblSaveLocation.AutoSize = true;
            LblSaveLocation.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            LblSaveLocation.Location = new Point(63, 140);
            LblSaveLocation.Name = "LblSaveLocation";
            LblSaveLocation.Size = new Size(109, 21);
            LblSaveLocation.TabIndex = 4;
            LblSaveLocation.Text = "Save Location:";
            // 
            // TxtExcelFile
            // 
            TxtExcelFile.Enabled = false;
            TxtExcelFile.Location = new Point(63, 39);
            TxtExcelFile.Name = "TxtExcelFile";
            TxtExcelFile.ScrollBars = ScrollBars.Both;
            TxtExcelFile.Size = new Size(591, 23);
            TxtExcelFile.TabIndex = 8;
            // 
            // TxtPpptFile
            // 
            TxtPpptFile.Location = new Point(63, 104);
            TxtPpptFile.Name = "TxtPpptFile";
            TxtPpptFile.ReadOnly = true;
            TxtPpptFile.Size = new Size(591, 23);
            TxtPpptFile.TabIndex = 9;
            // 
            // TxtSaveLocation
            // 
            TxtSaveLocation.Location = new Point(63, 167);
            TxtSaveLocation.Name = "TxtSaveLocation";
            TxtSaveLocation.Size = new Size(591, 23);
            TxtSaveLocation.TabIndex = 13;
            // 
            // LblMappingfileName
            // 
            LblMappingfileName.AutoSize = true;
            LblMappingfileName.Location = new Point(134, 175);
            LblMappingfileName.Name = "LblMappingfileName";
            LblMappingfileName.Size = new Size(0, 15);
            LblMappingfileName.TabIndex = 25;
            // 
            // CmdProcessPpt
            // 
            CmdProcessPpt.Font = new Font("Ebrima", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            CmdProcessPpt.Location = new Point(63, 205);
            CmdProcessPpt.Name = "CmdProcessPpt";
            CmdProcessPpt.Size = new Size(142, 38);
            CmdProcessPpt.TabIndex = 27;
            CmdProcessPpt.Text = "Create Offer";
            CmdProcessPpt.UseVisualStyleBackColor = true;
            CmdProcessPpt.Click += CmdProcessPpt_Click;
            // 
            // CmdLoadExcelFile
            // 
            CmdLoadExcelFile.Location = new Point(671, 39);
            CmdLoadExcelFile.Name = "CmdLoadExcelFile";
            CmdLoadExcelFile.Size = new Size(149, 23);
            CmdLoadExcelFile.TabIndex = 29;
            CmdLoadExcelFile.Text = "Load Excel File";
            CmdLoadExcelFile.UseVisualStyleBackColor = true;
            CmdLoadExcelFile.Click += CmdLoadExcelFile_Click;
            // 
            // CmdLoadPptFile
            // 
            CmdLoadPptFile.Location = new Point(671, 105);
            CmdLoadPptFile.Name = "CmdLoadPptFile";
            CmdLoadPptFile.Size = new Size(149, 23);
            CmdLoadPptFile.TabIndex = 30;
            CmdLoadPptFile.Text = "Load Powerpoint File";
            CmdLoadPptFile.UseVisualStyleBackColor = true;
            CmdLoadPptFile.Click += CmdLoadPptFile_Click;
            // 
            // ImgExcelLogo
            // 
            ImgExcelLogo.Image = Properties.Resources.Excel_Logo_40_40;
            ImgExcelLogo.Location = new Point(12, 22);
            ImgExcelLogo.MaximumSize = new Size(40, 40);
            ImgExcelLogo.MinimumSize = new Size(40, 40);
            ImgExcelLogo.Name = "ImgExcelLogo";
            ImgExcelLogo.Size = new Size(40, 40);
            ImgExcelLogo.TabIndex = 31;
            ImgExcelLogo.TabStop = false;
            // 
            // ImgPowerPointLogo
            // 
            ImgPowerPointLogo.Image = Properties.Resources.PowerPoint_Logo_40_40;
            ImgPowerPointLogo.Location = new Point(12, 87);
            ImgPowerPointLogo.MaximumSize = new Size(40, 40);
            ImgPowerPointLogo.MinimumSize = new Size(40, 40);
            ImgPowerPointLogo.Name = "ImgPowerPointLogo";
            ImgPowerPointLogo.Size = new Size(40, 40);
            ImgPowerPointLogo.TabIndex = 32;
            ImgPowerPointLogo.TabStop = false;
            // 
            // CmdSaveLocation
            // 
            CmdSaveLocation.Font = new Font("Ebrima", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            CmdSaveLocation.Location = new Point(671, 167);
            CmdSaveLocation.Name = "CmdSaveLocation";
            CmdSaveLocation.Size = new Size(149, 24);
            CmdSaveLocation.TabIndex = 33;
            CmdSaveLocation.Text = "Choose Save Location";
            CmdSaveLocation.UseVisualStyleBackColor = true;
            // 
            // ImgSaveDiskLogo
            // 
            ImgSaveDiskLogo.Image = Properties.Resources.FloppyDisk_Logo_50_50;
            ImgSaveDiskLogo.Location = new Point(7, 145);
            ImgSaveDiskLogo.MaximumSize = new Size(45, 45);
            ImgSaveDiskLogo.MinimumSize = new Size(45, 45);
            ImgSaveDiskLogo.Name = "ImgSaveDiskLogo";
            ImgSaveDiskLogo.Size = new Size(45, 45);
            ImgSaveDiskLogo.TabIndex = 34;
            ImgSaveDiskLogo.TabStop = false;
            // 
            // errorProvider1
            // 
            errorProvider1.ContainerControl = this;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(837, 256);
            Controls.Add(ImgSaveDiskLogo);
            Controls.Add(CmdSaveLocation);
            Controls.Add(ImgPowerPointLogo);
            Controls.Add(ImgExcelLogo);
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
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)ImgExcelLogo).EndInit();
            ((System.ComponentModel.ISupportInitialize)ImgPowerPointLogo).EndInit();
            ((System.ComponentModel.ISupportInitialize)ImgSaveDiskLogo).EndInit();
            ((System.ComponentModel.ISupportInitialize)errorProvider1).EndInit();
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
        private PictureBox ImgExcelLogo;
        private PictureBox ImgPowerPointLogo;
        private Button CmdSaveLocation;
        private PictureBox ImgSaveDiskLogo;
        private ErrorProvider errorProvider1;
    }
}
