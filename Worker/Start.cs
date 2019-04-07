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
        string htmlCode;
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
            MessageBox.Show("Start parse");
            string htmlCode;
            List<House> products = new List<House>();
            List<string> str = new List<string>();
            using (WebClient webClient = new WebClient())
            {
                // nastaveni ze webClient ma pouzit Windows Authentication
                // webClient.UseDefaultCredentials = true;
                webClient.Encoding = Encoding.UTF8;
                htmlCode = webClient.DownloadString(Link);
            }
            string patern = @"jss122"">(.*?)</span>";
            Match match = Regex.Match(htmlCode, patern);
            string strcount = match.Value;
            strcount = strcount.Replace(@"jss122"">", "");
            strcount = strcount.Replace(@" объявлений</span>", "");
            MessageBox.Show(strcount);
            patern = @"jss182""><a href=""(.*?)</div></div><a class=""jss91 jss65";
            Regex rgx = new Regex(patern);

            foreach (Match item in rgx.Matches(htmlCode))
            {
                str.Add(item.Value);
            }
            for (int i = 0; i < str.Count; i++)
            {
                House product = new House();
                products.Add(product);
            }
            for (int i = 0; i < str.Count; i++)
            {
                patern = @"jss206"">(.*?)</div>";
                match = Regex.Match(str[i], patern);
                string price = match.Value;
                price = price.Replace(@"jss206"">", "");
                price = price.Replace(@"</div>", "");
                products[i].Price = price;
                //Console.WriteLine(price);
                patern = @"class=""jss195"" title=""(.*?)""><span";
                match = Regex.Match(str[i], patern);
                string name = match.Value;
                name = name.Replace(@"class=""jss195"" title=""", "");
                name = name.Replace(@"""><span", "");
                products[i].Adress = name;
                //Console.WriteLine(name);
                patern = @"jss182""><a href=""(.*?)"" target";
                match = Regex.Match(str[i], patern);
                string link = match.Value;
                link = link.Replace(@"jss182""><a href=""", "");
                link = link.Replace(@""" class=""jss183"" target", "");
                products[i].Link = name;
                //Console.WriteLine(link);
                patern = @"<li class=""jss210"">(.*?)</li>";
                rgx = new Regex(patern);
                int index = 0;
                foreach (Match item in rgx.Matches(str[i]))
                {
                    string parametter = item.Value;
                    parametter = parametter.Replace(@"<li class=""jss210"">", "");
                    parametter = parametter.Replace(@"</li>", "");
                    parametter = parametter.Replace(@"<!-- -->", "");

                    if (index == 0)
                    {
                        products[i].CountOfRooms = parametter;
                        index++;
                    }
                    else
                    {
                        products[i].Metrazh = parametter;

                    }



                }
                //Console.WriteLine(products[i].Metrazh+" "+products[i].CountOfRooms);

            }
            patern = @"""updateTime"":""(.*?)"",""real";
            rgx = new Regex(patern);
            List<string> updates = new List<string>();
            foreach (Match item in rgx.Matches(htmlCode))
            {
                updates.Add(item.Value);
            }
            for (int i = 0; i < updates.Count; i++)
            {
                updates[i] = updates[i].Replace(@"""updateTime"":""", "");
                updates[i] = updates[i].Replace(@",""real", "");
                products[i].Updated = updates[i];
                //Console.WriteLine(updates[i]);
            }
            patern = @",""addTime"":""(.*?)"",""";
            rgx = new Regex(patern);
            List<string> addd = new List<string>();
            foreach (Match item in rgx.Matches(htmlCode))
            {
                addd.Add(item.Value);
            }
            for (int i = 0; i < addd.Count; i++)
            {
                addd[i] = addd[i].Replace(@",""addTime"":""", "");
                addd[i] = addd[i].Replace(@""",""", "");
                products[i].Updated = addd[i];
                Console.WriteLine(addd[i]);
            }
            int count = 0;

            int value = 0;
            foreach (char c in strcount)
            {
                if ((c >= '0') && (c <= '9'))
                {
                    value = value * 10 + (c - '0');
                }
            }
            value = value / 30;
            if (value / 30 > 1)
            {
                if(value > 100)
                {
                     count =100;
                }
                else
                {
                    count = value ;
                }
                for(int i=2; i<count-1;i++)
                {
                    List<House> products1 = new List<House>();
                    List<string> str1 = new List<string>();
                    string link = Link + "?page="+i.ToString();
                    using (WebClient webClient = new WebClient())
                    {
                        // nastaveni ze webClient ma pouzit Windows Authentication
                        // webClient.UseDefaultCredentials = true;
                        webClient.Encoding = Encoding.UTF8;
                        htmlCode = webClient.DownloadString(link);
                    }
                    #region a
                     patern = @"jss182""><a href=""(.*?)"" target";
                    rgx = new Regex(patern);

                    foreach (Match item in rgx.Matches(htmlCode))
                    {
                        str1.Add(item.Value);
                    }
                    for (int k = 0; k < str1.Count; k++)
                    {
                        House product = new House();
                        products1.Add(product);
                    }
                    for (int k = 0; k < str1.Count; k++)
                    {
                        //MessageBox.Show(k.ToString());
                        patern = @"jss206"">(.*?)</div>";
                        match = Regex.Match(str1[k], patern);
                        string price = match.Value;
                        price = price.Replace(@"jss206"">", "");
                        price = price.Replace(@"</div>", "");
                        products1[k].Price = price;
                        //Console.WriteLine(price);
                        patern = @"class=""jss195"" title=""(.*?)""><span";
                        match = Regex.Match(str1[k], patern);
                        string name = match.Value;
                        name = name.Replace(@"class=""jss195"" title=""", "");
                        name = name.Replace(@"""><span", "");
                        products1[k].Adress = name;
                        //Console.WriteLine(name);
                        patern = @"jss182""><a href=""(.*?)"" target";
                        match = Regex.Match(str1[k], patern);
                        link = match.Value;
                        link = link.Replace(@"jss182""><a href=""", "");
                        link = link.Replace(@""" class=""jss183"" target", "");
                        products1[k].Link = link;
                        //Console.WriteLine(link);
                        patern = @"<li class=""jss210"">(.*?)</li>";
                        rgx = new Regex(patern);
                        int index = 0;
                        foreach (Match item in rgx.Matches(str1[k]))
                        {
                            string parametter = item.Value;
                            parametter = parametter.Replace(@"<li class=""jss210"">", "");
                            parametter = parametter.Replace(@"</li>", "");
                            parametter = parametter.Replace(@"<!-- -->", "");

                            if (index == 0)
                            {
                                products[k].CountOfRooms = parametter;
                                index++;
                            }
                            else
                            {
                                products[k].Metrazh = parametter;

                            }



                        }
                        //Console.WriteLine(products[i].Metrazh+" "+products[i].CountOfRooms);

                    }
                    patern = @"""updateTime"":""(.*?)"",""real";
                    rgx = new Regex(patern);
                    List<string> updates1 = new List<string>();
                    foreach (Match item in rgx.Matches(htmlCode))
                    {
                        updates1.Add(item.Value);
                    }
                    for (int k = 0; k < updates.Count; k++)
                    {
                        updates1[k] = updates1[k].Replace(@"""updateTime"":""", "");
                        updates1[k] = updates1[k].Replace(@",""real", "");
                        products1[k].Updated = updates1[k];
                        //Console.WriteLine(updates[i]);
                    }
                    patern = @",""addTime"":""(.*?)"",""";
                    rgx = new Regex(patern);
                    List<string> addd1 = new List<string>();
                    foreach (Match item in rgx.Matches(htmlCode))
                    {
                        addd1.Add(item.Value);
                    }
                    for (int k = 0; k < addd.Count; k++)
                    {
                        addd1[k] = addd1[k].Replace(@",""addTime"":""", "");
                        addd1[k] = addd1[k].Replace(@""",""", "");
                        products1[k].Updated = addd1[k];
                        //Console.WriteLine(addd1[k]);
                    }
                    products.AddRange(products1);
                    #endregion
                }
            }
            MessageBox.Show(products.Count.ToString());
            MessageBox.Show("End parse");
        }


    
        private void button1_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    path = fbd.SelectedPath;
                }
            }
            filename = Microsoft.VisualBasic.Interaction.InputBox("Enter name of file (Without .excel) ", "Choosing name", "IAMSAMPLE");
            Thread InstanceCaller = new Thread(
            new ThreadStart(InstanceMethod));
            Link = textBox1.Text;
            InstanceCaller.Start();
        }
    }
}
