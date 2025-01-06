using Core.Domain.Models;
using Core.Interfaces;

namespace Offerten_Helper
{
    public partial class AutomatedOffer : Form
    {

        private readonly IExcelService _excelService;
        private readonly IPowerPointService _pptService;
        private readonly IMappingService _mappingService;

        private string _excelPath;
        private string _pptPath;
        public AutomatedOffer(IExcelService excelService,
            IPowerPointService pptService,
            IMappingService mappingService)
        {
            _excelService = excelService;
            _pptService = pptService;
            _mappingService = mappingService;
            InitializeComponent();
        }



        private void CmdProcessPpt_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_excelPath) || string.IsNullOrEmpty(_pptPath))
            {
                MessageBox.Show("Please select both Excel and PowerPoint files first.");
                return;
            }

            try
            {
                _excelService.LoadExcelFile(_excelPath);
                _pptService.LoadPowerPointFile(_pptPath);

                using var saveDialog = new SaveFileDialog
                {
                    Filter = "PowerPoint Files|*.pptx",
                    Title = "Save Offer As"
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    _mappingService.ApplyMapping(_excelPath, saveDialog.FileName);
                    MessageBox.Show("Offer created successfully!");
                }
            }
            catch (IOException ex) when (ex.Message.Contains("being used by another process"))
            {
                MessageBox.Show("Please close all Excel and PowerPoint files before creating the offer. One or more files are currently open.",
                    "File Access Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred while creating the offer: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

        }

        private void CmdLoadExcelFile_Click(object sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select Excel File"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _excelPath = dialog.FileName;

                try
                {
                    _excelService.LoadExcelFile(_excelPath);
                    TxtExcelFile.Text = _excelPath;
                }
                catch (IOException ex) when (ex.Message.Contains("being used by another process"))
                {
                    MessageBox.Show("Please close the Excel File before choosing.",
                        "File Access Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }


                if (TxtPpptFile.Text.Length != 0 && TxtExcelFile.Text.Length != 0)
                {
                    CmdProcessPpt.Enabled = true;
                    CmdCheckFiles.Enabled = true;

                    LblState.Text = "Ready to check files";

                }
            }
        }

        private void CmdLoadPptFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Powerpoint Files|*.pptx;*.pptx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                string[] mappings = new string[] { };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TxtPpptFile.Text = openFileDialog.FileName;
                    _pptPath = openFileDialog.FileName;
                }
            }


            if (TxtPpptFile.Text.Length != 0 && TxtExcelFile.Text.Length != 0)
            {
                CmdProcessPpt.Enabled = true;
                CmdCheckFiles.Enabled = true;

                LblState.Text = "Ready to check files";

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void CmdCheckFiles_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_excelPath) || string.IsNullOrEmpty(_pptPath))
            {
                MessageBox.Show("Please select both Excel and PowerPoint files first.");
                return;
            }

            try
            {
                _excelService.LoadExcelFile(_excelPath);
                _pptService.LoadPowerPointFile(_pptPath);

                var result = _mappingService.ValidateMapping(_excelPath, _pptPath);
                UpdateMissingNamesList(result);
            }
            catch (IOException ex) when (ex.Message.Contains("being used by another process"))
            {
                MessageBox.Show("Please close all Excel files before checking. One or more files are currently open.",
                    "File Access Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }


        //private void ValidateFiles()
        //{
        //    var result = UtilHelper.OfficeEditor.IsFieldNamesEqual(TxtPpptFile.Text, TxtExcelFile.Text);
        //    if (result.IsMappingMatching)
        //    {
        //        LblState.Text = "Success. No Missing Names found in PowerPoint";
        //        TxtStateDisplay.BackColor = Color.Green;

        //    }
        //    else
        //    {
        //        LblState.Text = "Some Names from Excel are missing in PowerPoint";
        //        TxtStateDisplay.BackColor = Color.Yellow;
        //    }

        //    LstPptMissingNames.Items.Clear();
        //    LstPptMissingNames.Items.AddRange(result.ExcelMissingMapping.ToArray());
        //}

        private void UpdateMissingNamesList(MappingResult result)
        {
            LstPptMissingNames.Items.Clear();
            foreach (var name in result.ExcelMissingMapping)
            {
                LstPptMissingNames.Items.Add(name);
            }
        }
    }
}
