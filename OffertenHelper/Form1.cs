namespace Offerten_Helper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void TxtPptTxtField_TextChanged(object sender, EventArgs e)
        {

        }

        private void CmdNewMapping_Click(object sender, EventArgs e)
        {
            LstMappings.Items.Add(TxtPptTxtField.Text);
        }

        private void CmdLoadMappingFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                string[] mappings = new string[] { };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    using (Stream fileStream = openFileDialog.OpenFile())
                    {
                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            string fileContent = reader.ReadToEnd();

                            fileContent.Trim();

                            mappings = fileContent.Split(";");

                            LstMappings.Items.Clear();

                            foreach (var mapping in mappings)
                            {
                                if (!mapping.Equals(""))
                                {
                                    LstMappings.Items.Add(mapping);
                                }
                            }
                        }
                        LblMappingfileName.Text = openFileDialog.FileName;
                    }
                }
            }
        }

        public string CreateOutputforMappingFile()
        {
            string mappings = "";

            foreach (var mapping in LstMappings.Items)
            {
                mappings += mapping.ToString() + ";";
            }

            return mappings;
        }

        private void CmdSaveMappingFile_Click(object sender, EventArgs e)
        {
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.Title = "Datei speichern";
                    saveFileDialog.DefaultExt = "txt";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        string mappings = CreateOutputforMappingFile();

                        try
                        {
                            File.WriteAllText(filePath, mappings);
                            MessageBox.Show("Datei wurde erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LblMappingfileName.Text = saveFileDialog.FileName;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Fehler beim Speichern der Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void CmdRemoveMapping_Click(object sender, EventArgs e)
        {
            if (LstMappings.SelectedItem != null)
            {
                LstMappings.Items.Remove(LstMappings.SelectedItem);
            }
        }

        private void CmdProcessPpt_Click(object sender, EventArgs e)
        {
            UtilHelper.UtilHelper.ProcessPowerpoint(TxtPpptFile.Text, TxtExcelFile.Text);
        }

        private void InsertValuesInPowerpoint(Dictionary<string, string> mappings, Dictionary<string, string> cellRefValue)
        {
            throw new NotImplementedException();
        }

        private Dictionary<string, string> CreateMappingDictionary()
        {
            Dictionary<string, string> mappings = new Dictionary<string, string>();

            foreach (var item in LstMappings.Items)
            {
                string[] mapping = item.ToString().Split(" ==> ");
                mappings.Add(mapping[0], mapping[1]);
            }

            return mappings;
        }

        private void CmdReadExcelValues_Click(object sender, EventArgs e)
        {
            List<string> mappingValues = new List<string>();

            foreach (var item in LstMappings.Items)
            {
                mappingValues.Add(item.ToString());
            }



            string excelPath = TxtExcelFile.Text;

            try
            {
                TxtTestField.Text = Controller.GetExcelData(excelPath, mappingValues);
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Warnung");

            }
        }

        private void CmdLoadExcelFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                string[] mappings = new string[] { };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TxtExcelFile.Text = openFileDialog.FileName;
                }
            }
        }

        private void CmdLoadPptFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                string[] mappings = new string[] { };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TxtPpptFile.Text = openFileDialog.FileName;
                }
            }
        }
    }
}
