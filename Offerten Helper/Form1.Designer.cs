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
            LblCells = new Label();
            LblTextfield = new Label();
            TxtExcelFile = new TextBox();
            TxtPpptFile = new TextBox();
            TxtSettingsFile = new TextBox();
            TxtCells = new TextBox();
            TxtPptTxtField = new TextBox();
            TxtSaveLocation = new TextBox();
            TxtNameNewFile = new TextBox();
            TxtMappings = new TextBox();
            TxtTestField = new TextBox();
            LblTesting = new Label();
            CmdTest = new Button();
            CmdNewMapping = new Button();
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
            // LblCells
            // 
            LblCells.AutoSize = true;
            LblCells.Location = new Point(230, 141);
            LblCells.Name = "LblCells";
            LblCells.Size = new Size(35, 15);
            LblCells.TabIndex = 6;
            LblCells.Text = "Cells:";
            // 
            // LblTextfield
            // 
            LblTextfield.AutoSize = true;
            LblTextfield.Location = new Point(525, 141);
            LblTextfield.Name = "LblTextfield";
            LblTextfield.Size = new Size(54, 15);
            LblTextfield.TabIndex = 7;
            LblTextfield.Text = "Textfield:";
            // 
            // TxtExcelFile
            // 
            TxtExcelFile.Location = new Point(10, 27);
            TxtExcelFile.Name = "TxtExcelFile";
            TxtExcelFile.Size = new Size(760, 23);
            TxtExcelFile.TabIndex = 8;
            // 
            // TxtPpptFile
            // 
            TxtPpptFile.Location = new Point(10, 71);
            TxtPpptFile.Name = "TxtPpptFile";
            TxtPpptFile.Size = new Size(760, 23);
            TxtPpptFile.TabIndex = 9;
            // 
            // TxtSettingsFile
            // 
            TxtSettingsFile.Location = new Point(10, 115);
            TxtSettingsFile.Name = "TxtSettingsFile";
            TxtSettingsFile.Size = new Size(760, 23);
            TxtSettingsFile.TabIndex = 10;
            // 
            // TxtCells
            // 
            TxtCells.Location = new Point(230, 159);
            TxtCells.Name = "TxtCells";
            TxtCells.Size = new Size(236, 23);
            TxtCells.TabIndex = 11;
            // 
            // TxtPptTxtField
            // 
            TxtPptTxtField.Location = new Point(525, 159);
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
            // TxtMappings
            // 
            TxtMappings.Location = new Point(12, 159);
            TxtMappings.Multiline = true;
            TxtMappings.Name = "TxtMappings";
            TxtMappings.Size = new Size(201, 194);
            TxtMappings.TabIndex = 15;
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
            // CmdTest
            // 
            CmdTest.Location = new Point(671, 476);
            CmdTest.Name = "CmdTest";
            CmdTest.Size = new Size(75, 23);
            CmdTest.TabIndex = 18;
            CmdTest.Text = "Excel lesen";
            CmdTest.UseVisualStyleBackColor = true;
            CmdTest.Click += CmdTest_Click;
            // 
            // CmdNewMapping
            // 
            CmdNewMapping.Location = new Point(230, 198);
            CmdNewMapping.Name = "CmdNewMapping";
            CmdNewMapping.Size = new Size(163, 23);
            CmdNewMapping.TabIndex = 19;
            CmdNewMapping.Text = "Create New Mapping";
            CmdNewMapping.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 777);
            Controls.Add(CmdNewMapping);
            Controls.Add(CmdTest);
            Controls.Add(LblTesting);
            Controls.Add(TxtTestField);
            Controls.Add(TxtMappings);
            Controls.Add(TxtNameNewFile);
            Controls.Add(TxtSaveLocation);
            Controls.Add(TxtPptTxtField);
            Controls.Add(TxtCells);
            Controls.Add(TxtSettingsFile);
            Controls.Add(TxtPpptFile);
            Controls.Add(TxtExcelFile);
            Controls.Add(LblTextfield);
            Controls.Add(LblCells);
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
        private Label LblCells;
        private Label LblTextfield;
        private TextBox TxtExcelFile;
        private TextBox TxtPpptFile;
        private TextBox TxtSettingsFile;
        private TextBox TxtCells;
        private TextBox TxtPptTxtField;
        private TextBox TxtSaveLocation;
        private TextBox TxtNameNewFile;
        private TextBox TxtMappings;
        private TextBox TxtTestField;
        private Label LblTesting;
        private Button CmdTest;
        private Button CmdNewMapping;
    }
}
