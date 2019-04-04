using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Worker
{
    public partial class Start : Form
    {
        private string path, filename;
        public Start()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    MessageBox.Show("Directory:" + fbd.SelectedPath);
                    path = fbd.SelectedPath;
                }
            }
            filename = Microsoft.VisualBasic.Interaction.InputBox("Enter name of file (Without .excel) ", "Choosing name", "IAMSAMPLE");
           
        }
    }
}
