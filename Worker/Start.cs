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
            List<House> houses = new List<House>();
            List<string> names = new List<string>();
            List<string> names2 = new List<string>();
            List<string> price = new List<string>();
            List<string> links = new List<string>();
            List<string> descriptions = new List<string>();
            List<string> parameters = new List<string>();
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
            strcount = strcount.Replace(@"jss122"">", "");
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
            patern = @"jss197"">(.*?)</span>";
            Regex rgx = new Regex(patern);
            
            foreach (Match item in rgx.Matches(htmlCode))
            {
                names.Add(item.Value);
            }
            for (int i = 0; i < names.Count; i++)
            {
                names[i] = names[i].Replace(@"jss197"">", "");
                names[i] = names[i].Replace(@"</span>", "");
            }
            patern = @"jss198"">, <!-- -->(.*?)</span>";
            rgx = new Regex(patern);

            foreach (Match item in rgx.Matches(htmlCode))
            {
                names2.Add(item.Value);
            }
            //MessageBox.Show(names2.Count.ToString());
            for (int i = 0; i < names2.Count; i++)
            {
                    names2[i] = names2[i].Replace(@"jss198"">, <!-- -->", "");
                    names2[i] = names2[i].Replace(@"</span>", "");
            }
            for(int i=0; i<names.Count; i++)
            {
                names[i] = names[i] + names2[i];
                //names[i] = names[i].Replace(@"<!-- -->", "");
            }
            //MessageBox.Show(names[0]);
            names2.Clear();
            patern = @"jss206"">(.*?)</div>";
            rgx = new Regex(patern);

            foreach (Match item in rgx.Matches(htmlCode))
            {
                price.Add(item.Value);
            }
            for (int i = 0; i < names.Count; i++)
            {
                price[i] = price[i].Replace(@"jss206"">", "");
                price[i] = price[i].Replace(@"</div>", "");
            }
            //MessageBox.Show(price[0]);
            //jss91 jss65 jss67 jss68 jss70 jss71 jss88 jss215" tabindex="0" role="button" href="                        " target
            //MessageBox.Show(price[0]);
            patern = @"jss91 jss65 jss67 jss68 jss70 jss71 jss88 jss215"" tabindex=""0"" role=""button"" href=(.*?)"" target";
            rgx = new Regex(patern);
            foreach (Match item in rgx.Matches(htmlCode))
            {
                links.Add(item.Value);
            }
            for (int i = 0; i < links.Count; i++)
            {
                links[i] = links[i].Replace(@"jss91 jss65 jss67 jss68 jss70 jss71 jss88 jss215"" tabindex=""0"" role=""button"" href=""", "");
                links[i] = links[i].Replace(@""" target", "");
                links[i] = "https://www.lun.ua" + links[i];
            }
            patern = @"<li class=""jss210"">(.*?)</li>";
            rgx = new Regex(patern);
            int index = 0;
            foreach (Match item in rgx.Matches(htmlCode))
            {
                if (index == 0)
                {
                    parameters.Add(item.Value);
                    index+=1;
                }
                else
                {
                    if(parameters[index-1].Contains("м²")==false && item.Value.Contains("м²") == true)
                    {
                        parameters.Add(item.Value);
                        index++;
                    }
                    else if(item.Value.Contains("м²") == false && parameters[index-1].Contains("м²") == true)
                    {
                        parameters.Add(item.Value);
                        index++;
                    }
                    


                }
                
                   
                
            }
            //int index = 0;
            for(int i=0; i<parameters.Count; i++)
            {
                        parameters[i] = parameters[i].Replace(@"<li class=""jss210"">", "");
                        parameters[i] = parameters[i].Replace(@"</li>", "");
                           
            }
            for(int i=0; i<parameters.Count; i++)
            {
                if(parameters[i].Contains(@"<!-- -->")==true)
                parameters[i] = parameters[i].Replace(@"<!-- -->", "");

            }
            //foreach (var item in parameters)
            //{
            //    MessageBox.Show(item);
            //}
            MessageBox.Show(parameters.Count.ToString());
            //MessageBox.Show(links.Count.ToString());
            //MessageBox.Show(names.Count.ToString());
            //MessageBox.Show(price.Count.ToString());

            //MessageBox.Show(links[0]);
            countofpages = value / 30;
            if (countofpages > 100)
            {
                for (int i = 2; i <= 100; i++)
                {

                }
            }
            else
            {
                countofpages = countofpages + 1;
                for (int i = 2; i <= countofpages; i++)
                {

                }
            }
            //MessageBox.Show(countofpages.ToString()) ;
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
