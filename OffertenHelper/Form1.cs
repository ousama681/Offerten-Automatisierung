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
            UtilHelper.UtilHelper.ProcessPowerpoint(TxtPpptFile.Text, TxtExcelFile.Text);
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
