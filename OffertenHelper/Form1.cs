using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

namespace Offerten_Helper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void CmdProcessPpt_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Powerpoint Files|*.pptx;*.pptx";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            string pptFile = TxtPpptFile.Text.Split("\\").Last();

            string toReplace = "_Processed.pptx";

            string newPptFile = pptFile.Replace(".pptx", toReplace);


            saveFileDialog.FileName = newPptFile;


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = UtilHelper.UtilHelper.ProcessPowerpoint(TxtPpptFile.Text, TxtExcelFile.Text, saveFileDialog.FileName);

                DialogResult dialogResult = MessageBox.Show("Offer created. Would you like to review the offer?", "Success", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dialogResult == DialogResult.Yes)
                {
                    try
                    {
                        if (File.Exists(filePath))
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = filePath,
                                UseShellExecute = true
                            });
                        }
                        else
                        {
                            MessageBox.Show("Die angegebene Datei existiert nicht.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Fehler beim Öffnen der Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void CmdLoadExcelFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel Files|*.xlsx;";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                string[] mappings = new string[] { };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TxtExcelFile.Text = openFileDialog.FileName;
                }
            }


            if (TxtPpptFile.Text.Length != 0 && TxtExcelFile.Text.Length != 0)
            {
                CmdProcessPpt.Enabled = true;
                CmdCheckFiles.Enabled = true;

                TxtStateDisplay.BackColor = Color.AliceBlue;
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
                }
            }


            if (TxtPpptFile.Text.Length != 0 && TxtExcelFile.Text.Length != 0)
            {
                CmdProcessPpt.Enabled = true;
                CmdCheckFiles.Enabled = true;


            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void CmdCheckFiles_Click(object sender, EventArgs e)
        {
            ValidateFiles();

            MessageBox.Show("Files have been checked.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ValidateFiles()
        {
            var result = UtilHelper.UtilHelper.IsDefinedNamesEqual(TxtPpptFile.Text, TxtExcelFile.Text);
            if (result.IsMappingMatching)
            {
                TxtState.Text = "No complications found.";
                TxtStateDisplay.BackColor = Color.Green;

            }
            else
            {
                TxtState.Text = "Warning: Not all Defined Names match.";
                TxtStateDisplay.BackColor = Color.Yellow;
            }

            LstPptMissingNames.Items.Clear();
            LstPptMissingNames.Items.AddRange(result.ExcelMissingMapping.ToArray());
        }
    }
}
