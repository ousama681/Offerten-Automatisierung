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
            string filePath = UtilHelper.UtilHelper.ProcessPowerpoint(TxtPpptFile.Text, TxtExcelFile.Text);

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
                    MessageBox.Show($"Fehler beim �ffnen der Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void CmdSaveLocation_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.pptx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                string[] mappings = new string[] { };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TxtSaveLocation.Text = openFileDialog.FileName;
                }
            }
        }
    }
}
