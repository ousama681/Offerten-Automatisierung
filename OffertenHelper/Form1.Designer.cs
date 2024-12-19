﻿namespace Offerten_Helper
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
            LblSaveLocation = new Label();
            TxtExcelFile = new TextBox();
            TxtPpptFile = new TextBox();
            TxtSaveLocation = new TextBox();
            LblMappingfileName = new Label();
            CmdProcessPpt = new Button();
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
            // LblSaveLocation
            // 
            LblSaveLocation.AutoSize = true;
            LblSaveLocation.Location = new Point(10, 106);
            LblSaveLocation.Name = "LblSaveLocation";
            LblSaveLocation.Size = new Size(83, 15);
            LblSaveLocation.TabIndex = 4;
            LblSaveLocation.Text = "Save Location:";
            // 
            // TxtExcelFile
            // 
            TxtExcelFile.Enabled = false;
            TxtExcelFile.Location = new Point(10, 27);
            TxtExcelFile.Name = "TxtExcelFile";
            TxtExcelFile.ScrollBars = ScrollBars.Both;
            TxtExcelFile.Size = new Size(875, 23);
            TxtExcelFile.TabIndex = 8;
            // 
            // TxtPpptFile
            // 
            TxtPpptFile.Location = new Point(10, 71);
            TxtPpptFile.Name = "TxtPpptFile";
            TxtPpptFile.ReadOnly = true;
            TxtPpptFile.Size = new Size(875, 23);
            TxtPpptFile.TabIndex = 9;
            // 
            // TxtSaveLocation
            // 
            TxtSaveLocation.Location = new Point(10, 124);
            TxtSaveLocation.Name = "TxtSaveLocation";
            TxtSaveLocation.Size = new Size(875, 23);
            TxtSaveLocation.TabIndex = 13;
            // 
            // LblMappingfileName
            // 
            LblMappingfileName.AutoSize = true;
            LblMappingfileName.Location = new Point(81, 141);
            LblMappingfileName.Name = "LblMappingfileName";
            LblMappingfileName.Size = new Size(0, 15);
            LblMappingfileName.TabIndex = 25;
            // 
            // CmdProcessPpt
            // 
            CmdProcessPpt.Location = new Point(10, 159);
            CmdProcessPpt.Name = "CmdProcessPpt";
            CmdProcessPpt.Size = new Size(163, 23);
            CmdProcessPpt.TabIndex = 27;
            CmdProcessPpt.Text = "Process Powerpoint";
            CmdProcessPpt.UseVisualStyleBackColor = true;
            CmdProcessPpt.Click += CmdProcessPpt_Click;
            // 
            // CmdLoadExcelFile
            // 
            CmdLoadExcelFile.Location = new Point(891, 27);
            CmdLoadExcelFile.Name = "CmdLoadExcelFile";
            CmdLoadExcelFile.Size = new Size(149, 23);
            CmdLoadExcelFile.TabIndex = 29;
            CmdLoadExcelFile.Text = "Load Excel File";
            CmdLoadExcelFile.UseVisualStyleBackColor = true;
            CmdLoadExcelFile.Click += CmdLoadExcelFile_Click;
            // 
            // CmdLoadPptFile
            // 
            CmdLoadPptFile.Location = new Point(891, 71);
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
            ClientSize = new Size(1052, 205);
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
    }
}
