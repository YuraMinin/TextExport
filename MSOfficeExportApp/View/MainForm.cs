using MSOfficeExportApp.Service;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MSOfficeExportApp
{
    public partial class MainForm : Form
    {
        public delegate void Export(String dir, String text);

        private String _dir;
        private Export _exportText;        

        public MainForm()
        {
            InitializeComponent();
            
            textBox.Enabled = false;
            exportBtn.Enabled = false;
        }

        public void registerExportDelegate(Export exportText)
        {
            _exportText += exportText;
        }

        private async void Export_Click(object sender, EventArgs e)
        {
            String text = textBox.Text;
            await Task.Run(() =>
           {
               _exportText?.Invoke(_dir, text);
           });
            
        }

        private void Choose_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                _dir = folder.SelectedPath;
                textBox.Enabled = true;
                exportBtn.Enabled = true;
                l_directory.Text = "Directory: " + _dir;
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
