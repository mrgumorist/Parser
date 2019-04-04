using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Worker
{
    public partial class Start : Form
    {
        public static string path, filename, Link;
       
        public Start()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        public static void InstanceMethod()
        {
            string htmlCode;
            double countofpages;
            MessageBox.Show("We are start our parsing");
            using (WebClient client = new WebClient()) // WebClient class inherits IDisposable
            {
                client.Encoding = Encoding.UTF8;
                 htmlCode = client.DownloadString(Link);
                File.WriteAllText("first.txt", htmlCode);
            }
            //< span class="jss122">867 объявлений</span>
            string patern = @"jss122"">(.*?)</span>";
            Match match = Regex.Match(htmlCode, patern);
            string strcount = match.Value;
            strcount=strcount.Replace(@"jss122"">", "");
           // MessageBox.Show(strcount);
            int value = 0;
            foreach (char c in strcount)
            {
                if ((c >= '0') && (c <= '9'))
                {
                    value = value * 10 + (c - '0');
                }
            }
            //MessageBox.Show(value.ToString()) ;
            countofpages = value / 30;
            MessageBox.Show(countofpages.ToString()) ;
            MessageBox.Show("We are end our parsing");
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    //MessageBox.Show("Directory:" + fbd.SelectedPath);
                    path = fbd.SelectedPath;
                }
            }
            filename = Microsoft.VisualBasic.Interaction.InputBox("Enter name of file (Without .excel) ", "Choosing name", "IAMSAMPLE");
            Thread InstanceCaller = new Thread(
            new ThreadStart(InstanceMethod));
            Link = textBox1.Text;

            // Start the thread.
            InstanceCaller.Start();
        }
    }
}
