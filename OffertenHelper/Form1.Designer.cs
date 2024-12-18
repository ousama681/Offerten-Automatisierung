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
            LblExcelFile = new Label();
            LblPptFile = new Label();
            LblSettings = new Label();
            LblMappings = new Label();
            LblSaveLocation = new Label();
            LblNewName = new Label();
            LblTextfield = new Label();
            TxtExcelFile = new TextBox();
            TxtPpptFile = new TextBox();
            TxtSettingsFile = new TextBox();
            TxtPptTxtField = new TextBox();
            TxtSaveLocation = new TextBox();
            TxtNameNewFile = new TextBox();
            TxtTestField = new TextBox();
            LblTesting = new Label();
            CmdNewMapping = new Button();
            LstMappings = new ListBox();
            CmdSaveMappingFile = new Button();
            CmdLoadMappingFile = new Button();
            LblMappingfileName = new Label();
            CmdRemoveMapping = new Button();
            CmdProcessPpt = new Button();
            CmdReadExcelValues = new Button();
            CmdLoadExcelFile = new Button();
            CmdLoadPptFile = new Button();
            SuspendLayout();
            // 
            // LblExcelFile
            // 
            LblExcelFile.AutoSize = true;
            LblExcelFile.Location = new Point(10, 9);
            LblExcelFile.Name = "LblExcelFile";
            LblExcelFile.Size = new Size(55, 15);
            LblExcelFile.TabIndex = 0;
            LblExcelFile.Text = "ExcelFile:";
            // 
            // LblPptFile
            // 
            LblPptFile.AutoSize = true;
            LblPptFile.Location = new Point(10, 53);
            LblPptFile.Name = "LblPptFile";
            LblPptFile.Size = new Size(92, 15);
            LblPptFile.TabIndex = 1;
            LblPptFile.Text = "PowerPoint File:";
            // 
            // LblSettings
            // 
            LblSettings.AutoSize = true;
            LblSettings.Location = new Point(13, 97);
            LblSettings.Name = "LblSettings";
            LblSettings.Size = new Size(52, 15);
            LblSettings.TabIndex = 2;
            LblSettings.Text = "Settings:";
            // 
            // LblMappings
            // 
            LblMappings.AutoSize = true;
            LblMappings.Location = new Point(12, 141);
            LblMappings.Name = "LblMappings";
            LblMappings.Size = new Size(63, 15);
            LblMappings.TabIndex = 3;
            LblMappings.Text = "Mappings:";
            // 
            // LblSaveLocation
            // 
            LblSaveLocation.AutoSize = true;
            LblSaveLocation.Location = new Point(10, 356);
            LblSaveLocation.Name = "LblSaveLocation";
            LblSaveLocation.Size = new Size(83, 15);
            LblSaveLocation.TabIndex = 4;
            LblSaveLocation.Text = "Save Location:";
            // 
            // LblNewName
            // 
            LblNewName.AutoSize = true;
            LblNewName.Location = new Point(10, 400);
            LblNewName.Name = "LblNewName";
            LblNewName.Size = new Size(145, 15);
            LblNewName.TabIndex = 5;
            LblNewName.Text = "Name of new PowerPoint:";
            // 
            // LblTextfield
            // 
            LblTextfield.AutoSize = true;
            LblTextfield.Location = new Point(231, 141);
            LblTextfield.Name = "LblTextfield";
            LblTextfield.Size = new Size(54, 15);
            LblTextfield.TabIndex = 7;
            LblTextfield.Text = "Textfield:";
            // 
            // TxtExcelFile
            // 
            TxtExcelFile.Enabled = false;
            TxtExcelFile.Location = new Point(10, 27);
            TxtExcelFile.Name = "TxtExcelFile";
            TxtExcelFile.ScrollBars = ScrollBars.Both;
            TxtExcelFile.Size = new Size(605, 23);
            TxtExcelFile.TabIndex = 8;
            // 
            // TxtPpptFile
            // 
            TxtPpptFile.Location = new Point(10, 71);
            TxtPpptFile.Name = "TxtPpptFile";
            TxtPpptFile.ReadOnly = true;
            TxtPpptFile.Size = new Size(605, 23);
            TxtPpptFile.TabIndex = 9;
            // 
            // TxtSettingsFile
            // 
            TxtSettingsFile.Location = new Point(10, 115);
            TxtSettingsFile.Name = "TxtSettingsFile";
            TxtSettingsFile.Size = new Size(760, 23);
            TxtSettingsFile.TabIndex = 10;
            // 
            // TxtPptTxtField
            // 
            TxtPptTxtField.Location = new Point(231, 159);
            TxtPptTxtField.Name = "TxtPptTxtField";
            TxtPptTxtField.Size = new Size(245, 23);
            TxtPptTxtField.TabIndex = 12;
            TxtPptTxtField.TextChanged += TxtPptTxtField_TextChanged;
            // 
            // TxtSaveLocation
            // 
            TxtSaveLocation.Location = new Point(10, 374);
            TxtSaveLocation.Name = "TxtSaveLocation";
            TxtSaveLocation.Size = new Size(760, 23);
            TxtSaveLocation.TabIndex = 13;
            // 
            // TxtNameNewFile
            // 
            TxtNameNewFile.Location = new Point(10, 418);
            TxtNameNewFile.Name = "TxtNameNewFile";
            TxtNameNewFile.Size = new Size(760, 23);
            TxtNameNewFile.TabIndex = 14;
            // 
            // TxtTestField
            // 
            TxtTestField.Location = new Point(10, 473);
            TxtTestField.Multiline = true;
            TxtTestField.Name = "TxtTestField";
            TxtTestField.Size = new Size(650, 266);
            TxtTestField.TabIndex = 16;
            // 
            // LblTesting
            // 
            LblTesting.AutoSize = true;
            LblTesting.Location = new Point(10, 455);
            LblTesting.Name = "LblTesting";
            LblTesting.Size = new Size(67, 15);
            LblTesting.TabIndex = 17;
            LblTesting.Text = "For Testing:";
            // 
            // CmdNewMapping
            // 
            CmdNewMapping.Location = new Point(230, 188);
            CmdNewMapping.Name = "CmdNewMapping";
            CmdNewMapping.Size = new Size(163, 23);
            CmdNewMapping.TabIndex = 19;
            CmdNewMapping.Text = "Add New Mapping";
            CmdNewMapping.UseVisualStyleBackColor = true;
            CmdNewMapping.Click += CmdNewMapping_Click;
            // 
            // LstMappings
            // 
            LstMappings.FormattingEnabled = true;
            LstMappings.ItemHeight = 15;
            LstMappings.Items.AddRange(new object[] { "CheeseVatType" });
            LstMappings.Location = new Point(10, 159);
            LstMappings.Name = "LstMappings";
            LstMappings.Size = new Size(196, 184);
            LstMappings.TabIndex = 20;
            // 
            // CmdSaveMappingFile
            // 
            CmdSaveMappingFile.Location = new Point(231, 291);
            CmdSaveMappingFile.Name = "CmdSaveMappingFile";
            CmdSaveMappingFile.Size = new Size(142, 23);
            CmdSaveMappingFile.TabIndex = 23;
            CmdSaveMappingFile.Text = "Save Mappings";
            CmdSaveMappingFile.UseVisualStyleBackColor = true;
            CmdSaveMappingFile.Click += CmdSaveMappingFile_Click;
            // 
            // CmdLoadMappingFile
            // 
            CmdLoadMappingFile.Location = new Point(525, 291);
            CmdLoadMappingFile.Name = "CmdLoadMappingFile";
            CmdLoadMappingFile.Size = new Size(149, 23);
            CmdLoadMappingFile.TabIndex = 24;
            CmdLoadMappingFile.Text = "Load Mappings";
            CmdLoadMappingFile.UseVisualStyleBackColor = true;
            CmdLoadMappingFile.Click += CmdLoadMappingFile_Click;
            // 
            // LblMappingfileName
            // 
            LblMappingfileName.AutoSize = true;
            LblMappingfileName.Location = new Point(81, 141);
            LblMappingfileName.Name = "LblMappingfileName";
            LblMappingfileName.Size = new Size(0, 15);
            LblMappingfileName.TabIndex = 25;
            // 
            // CmdRemoveMapping
            // 
            CmdRemoveMapping.Location = new Point(230, 217);
            CmdRemoveMapping.Name = "CmdRemoveMapping";
            CmdRemoveMapping.Size = new Size(163, 23);
            CmdRemoveMapping.TabIndex = 26;
            CmdRemoveMapping.Text = "Remove Mapping";
            CmdRemoveMapping.UseVisualStyleBackColor = true;
            CmdRemoveMapping.Click += CmdRemoveMapping_Click;
            // 
            // CmdProcessPpt
            // 
            CmdProcessPpt.Location = new Point(525, 217);
            CmdProcessPpt.Name = "CmdProcessPpt";
            CmdProcessPpt.Size = new Size(163, 23);
            CmdProcessPpt.TabIndex = 27;
            CmdProcessPpt.Text = "Process Powerpoint";
            CmdProcessPpt.UseVisualStyleBackColor = true;
            CmdProcessPpt.Click += CmdProcessPpt_Click;
            // 
            // CmdReadExcelValues
            // 
            CmdReadExcelValues.Location = new Point(666, 473);
            CmdReadExcelValues.Name = "CmdReadExcelValues";
            CmdReadExcelValues.Size = new Size(75, 23);
            CmdReadExcelValues.TabIndex = 28;
            CmdReadExcelValues.Text = "Excel lesen";
            CmdReadExcelValues.UseVisualStyleBackColor = true;
            CmdReadExcelValues.Click += CmdReadExcelValues_Click;
            // 
            // CmdLoadExcelFile
            // 
            CmdLoadExcelFile.Location = new Point(621, 27);
            CmdLoadExcelFile.Name = "CmdLoadExcelFile";
            CmdLoadExcelFile.Size = new Size(149, 23);
            CmdLoadExcelFile.TabIndex = 29;
            CmdLoadExcelFile.Text = "Load Excel File";
            CmdLoadExcelFile.UseVisualStyleBackColor = true;
            CmdLoadExcelFile.Click += CmdLoadExcelFile_Click;
            // 
            // CmdLoadPptFile
            // 
            CmdLoadPptFile.Location = new Point(621, 70);
            CmdLoadPptFile.Name = "CmdLoadPptFile";
            CmdLoadPptFile.Size = new Size(149, 23);
            CmdLoadPptFile.TabIndex = 30;
            CmdLoadPptFile.Text = "Load Powerpoint File";
            CmdLoadPptFile.UseVisualStyleBackColor = true;
            CmdLoadPptFile.Click += CmdLoadPptFile_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 777);
            Controls.Add(CmdLoadPptFile);
            Controls.Add(CmdLoadExcelFile);
            Controls.Add(CmdReadExcelValues);
            Controls.Add(CmdProcessPpt);
            Controls.Add(CmdRemoveMapping);
            Controls.Add(LblMappingfileName);
            Controls.Add(CmdLoadMappingFile);
            Controls.Add(CmdSaveMappingFile);
            Controls.Add(LstMappings);
            Controls.Add(CmdNewMapping);
            Controls.Add(LblTesting);
            Controls.Add(TxtTestField);
            Controls.Add(TxtNameNewFile);
            Controls.Add(TxtSaveLocation);
            Controls.Add(TxtPptTxtField);
            Controls.Add(TxtSettingsFile);
            Controls.Add(TxtPpptFile);
            Controls.Add(TxtExcelFile);
            Controls.Add(LblTextfield);
            Controls.Add(LblNewName);
            Controls.Add(LblSaveLocation);
            Controls.Add(LblMappings);
            Controls.Add(LblSettings);
            Controls.Add(LblPptFile);
            Controls.Add(LblExcelFile);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label LblExcelFile;
        private Label LblPptFile;
        private Label LblSettings;
        private Label LblMappings;
        private Label LblSaveLocation;
        private Label LblNewName;
        private Label LblTextfield;
        private TextBox TxtExcelFile;
        private TextBox TxtPpptFile;
        private TextBox TxtSettingsFile;
        private TextBox TxtPptTxtField;
        private TextBox TxtSaveLocation;
        private TextBox TxtNameNewFile;
        private TextBox TxtTestField;
        private Label LblTesting;
        private Button CmdNewMapping;
        private ListBox LstMappings;
        private Button CmdSaveMappingFile;
        private Button CmdLoadMappingFile;
        private Label LblMappingfileName;
        private Button CmdRemoveMapping;
        private Button CmdProcessPpt;
        private Button CmdReadExcelValues;
        private Button CmdLoadExcelFile;
        private Button CmdLoadPptFile;
    }
}
