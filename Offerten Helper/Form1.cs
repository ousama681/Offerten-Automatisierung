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

        private void CmdTest_Click(object sender, EventArgs e)
        {
            string excelPath = TxtExcelFile.Text;
            TxtTestField.Text = Controller.GetExcelData(excelPath, true);
        }
    }
}
