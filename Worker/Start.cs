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
using OfficeOpenXml;
using OfficeOpenXml.Style;

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

        private void button2_Click(object sender, EventArgs e)
        {
            //using (var fbd = new FolderBrowserDialog())
            //{
            //    DialogResult result = fbd.ShowDialog();

            //    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            //    {
            //        path = fbd.SelectedPath;
            //    }
            //}
           // filename = Microsoft.VisualBasic.Interaction.InputBox("Enter name of file (Without .excel) ", "Choosing name", "IAMSAMPLE");
            Thread InstanceCaller = new Thread(
            new ThreadStart(Method));
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files(*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName;
            }
            Link = textBox1.Text;
            InstanceCaller.Start();
        }
        public static string filename1;
        public static void Method()
        {
            MessageBox.Show("Start parsing");
            List<string> fromfile = new List<string>();
            var fi = new FileInfo(filename);
           
           
            using (var package = new ExcelPackage(fi))
                    {
                        var workbook = package.Workbook;
                        var worksheet = workbook.Worksheets.Last();
                        //TODO a
                        int i = 2;
                        var rowCnt = worksheet.Dimension.End.Row;
                        
                        while (i!=rowCnt+1)
                        {
                            fromfile.Add(worksheet.Cells[i, 1].Value.ToString());
                            i++;
                        }

                //MessageBox.Show(fromfile.Count.ToString());
                //MessageBox.Show(products.Count.ToString());
                package.Dispose();
            }
            //!list.Exists(x => x.ID == 1)
            #region a
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
            //MessageBox.Show(strcount);
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
                patern = @"""jss195"" title=""(.*?)><div class";
                match = Regex.Match(str[i], patern);
                string name = match.Value;
                name = name.Replace(@"""jss195"" title=""", "");
                name = name.Replace(@"><div class", "");
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
                //MessageBox.Show(products[i].Adress);
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
                products[i].Created = addd[i];
                //Console.WriteLine(addd[i]);
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
                if (value > 100)
                {
                    count = 100;
                }
                else
                {
                    count = value;
                }
                for (int k = 2; k < count - 1; k++)
                {
                    List<House> products1 = new List<House>();
                    List<string> str1 = new List<string>();
                    string link = Link + "?page=" + k.ToString();
                    using (WebClient webClient = new WebClient())
                    {
                        // nastaveni ze webClient ma pouzit Windows Authentication
                        // webClient.UseDefaultCredentials = true;
                        webClient.Encoding = Encoding.UTF8;
                        htmlCode = webClient.DownloadString(link);
                    }
                  
                    patern = @"jss182""><a href=""(.*?)"" target";
                    rgx = new Regex(patern);

                    foreach (Match item in rgx.Matches(htmlCode))
                    {
                        str1.Add(item.Value);
                    }
                    for (int i = 0; i < str1.Count; i++)
                    {
                        patern = @"jss206"">(.*?)</div>";
                        match = Regex.Match(str1[i], patern);
                        string price = match.Value;
                        price = price.Replace(@"jss206"">", "");
                        price = price.Replace(@"</div>", "");
                        products1[i].Price = price;
                        //Console.WriteLine(price);
                        patern = @"""jss195"" title=""(.*?)><div class";
                        match = Regex.Match(str1[i], patern);
                        string name = match.Value;
                        name = name.Replace(@"""jss195"" title=""", "");
                        name = name.Replace(@"><div class", "");
                        products1[i].Adress = name;
                        //Console.WriteLine(name);
                        patern = @"jss182""><a href=""(.*?)"" target";
                        match = Regex.Match(str1[i], patern);
                        link = match.Value;
                        link = link.Replace(@"jss182""><a href=""", "");
                        link = link.Replace(@""" class=""jss183"" target", "");
                        products1[i].Link = name;
                        //Console.WriteLine(link);
                        patern = @"<li class=""jss210"">(.*?)</li>";
                        rgx = new Regex(patern);
                        int index = 0;
                        foreach (Match item in rgx.Matches(str1[i]))
                        {
                            string parametter = item.Value;
                            parametter = parametter.Replace(@"<li class=""jss210"">", "");
                            parametter = parametter.Replace(@"</li>", "");
                            parametter = parametter.Replace(@"<!-- -->", "");

                            if (index == 0)
                            {
                                products1[i].CountOfRooms = parametter;
                                index++;
                            }
                            else
                            {
                                products1[i].Metrazh = parametter;

                            }



                        }
                    }
                    patern = @"""updateTime"":""(.*?)"",""real";
                    rgx = new Regex(patern);
                    List<string> updates1 = new List<string>();
                    foreach (Match item in rgx.Matches(htmlCode))
                    {
                        updates1.Add(item.Value);
                    }
                    for (int l = 0; l < updates1.Count; l++)
                    {
                        updates1[l] = updates1[l].Replace(@"""updateTime"":""", "");
                        updates1[l] = updates1[l].Replace(@",""real", "");
                        products1[l].Updated = updates1[l];
                        //Console.WriteLine(updates[i]);
                    }
                    patern = @",""addTime"":""(.*?)"",""";
                    rgx = new Regex(patern);
                    List<string> addd1 = new List<string>();
                    foreach (Match item in rgx.Matches(htmlCode))
                    {
                        addd1.Add(item.Value);
                    }
                    for (int l = 0; l < addd1.Count; l++)
                    {
                        addd1[l] = addd1[l].Replace(@",""addTime"":""", "");
                        addd1[l] = addd1[l].Replace(@""",""", "");
                        products1[l].Created = addd1[l];
                        //Console.WriteLine(addd[i]);
                    }

                    products.AddRange(products1);
                    
                }
            }
            #endregion
            List<string> productsNew = new List<string>();
            foreach (var item in products)
            {
                productsNew.Add(item.Adress);
            }
            List<string> isnue = new List<string>();
            for(int i=0; i<products.Count; i++)
            {
                isnue.Add("Old");
            }
            for(int i=0; i< productsNew.Count; i++)
            {
                if(!fromfile.Contains(productsNew[i]))
                {
                    isnue[i] = "New";
                }
            }
            using (var package = new ExcelPackage(fi))
            {
                Random random = new Random();
                int rand = random.Next(10000000);
                var workbook = package.Workbook;
                package.Workbook.Worksheets.Add("Worksheetlast" + rand);
                var excelWorksheet = package.Workbook.Worksheets["Worksheetlast" + rand];
                //TODO a
                //int i = 2;
                //var rowCnt = excelWorksheet.Dimension.End.Row;

                //while (i != rowCnt + 1)
                //{
                //    fromfile.Add(excelWorksheet.Cells[i, 1].Value.ToString());
                //    i++;
                //}
                excelWorksheet.Cells[1, 1].Value = "Adress";
                excelWorksheet.Cells[1, 2].Value = "Price";
                excelWorksheet.Cells[1, 3].Value = "CountOfRooms";
                excelWorksheet.Cells[1, 4].Value = "Metrazh";
                excelWorksheet.Cells[1, 5].Value = "Link";
                excelWorksheet.Cells[1, 6].Value = "Created";
                excelWorksheet.Cells[1, 7].Value = "Updated";
                excelWorksheet.Cells[1, 8].Value = "ISNEW";
                for (int i = 0; i < products.Count; i++)
                {
                    excelWorksheet.Cells[i + 2, 1].Value = products[i].Adress;
                    excelWorksheet.Cells[i + 2, 2].Value = products[i].Price;
                    excelWorksheet.Cells[i + 2, 3].Value = products[i].CountOfRooms;
                    excelWorksheet.Cells[i + 2, 4].Value = products[i].Metrazh;
                    excelWorksheet.Cells[i + 2, 5].Value = products[i].Link;
                    excelWorksheet.Cells[i + 2, 6].Value = products[i].Created;
                    excelWorksheet.Cells[i + 2, 7].Value = products[i].Updated;
                    ExcelStyles excelStyles;
                     if(isnue[i]=="New")
                     {
                        Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        excelWorksheet.Cells[i+2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 1].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 2].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 3].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 5].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 6].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 7].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    }
                     else
                     {
                        Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#ffff00");
                        excelWorksheet.Cells[i + 2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 1].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 2].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 3].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 5].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 6].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        excelWorksheet.Cells[i + 2, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelWorksheet.Cells[i + 2, 7].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    }
                    excelWorksheet.Cells[i + 2, 8].Value = isnue[i];
                    // MessageBox.Show(products[i].Adress);
                }
                package.Save();
                package.Dispose();
                //MessageBox.Show(fromfile.Count.ToString());
                //MessageBox.Show(products.Count.ToString());
            }
            MessageBox.Show("End parsing");



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
            //MessageBox.Show(strcount);
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
                patern = @"""jss195"" title=""(.*?)><div class";
                match = Regex.Match(str[i], patern);
                string name = match.Value;
                name = name.Replace(@"""jss195"" title=""", "");
                name = name.Replace(@"><div class", "");
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
                //MessageBox.Show(products[i].Adress);
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
                products[i].Created = addd[i];
                //Console.WriteLine(addd[i]);
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
                for(int k=2; k<count-1;k++)
                {
                    List<House> products1 = new List<House>();
                    List<string> str1 = new List<string>();
                    string link = Link + "?page="+k.ToString();
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
                    for (int i = 0; i < str1.Count; i++)
                    {
                        patern = @"jss206"">(.*?)</div>";
                        match = Regex.Match(str1[i], patern);
                        string price = match.Value;
                        price = price.Replace(@"jss206"">", "");
                        price = price.Replace(@"</div>", "");
                        products1[i].Price = price;
                        //Console.WriteLine(price);
                        patern = @"""jss195"" title=""(.*?)><div class";
                        match = Regex.Match(str1[i], patern);
                        string name = match.Value;
                        name = name.Replace(@"""jss195"" title=""", "");
                        name = name.Replace(@"><div class", "");
                        products1[i].Adress = name;
                        //Console.WriteLine(name);
                        patern = @"jss182""><a href=""(.*?)"" target";
                        match = Regex.Match(str1[i], patern);
                        link = match.Value;
                        link = link.Replace(@"jss182""><a href=""", "");
                        link = link.Replace(@""" class=""jss183"" target", "");
                        products1[i].Link = name;
                        //Console.WriteLine(link);
                        patern = @"<li class=""jss210"">(.*?)</li>";
                        rgx = new Regex(patern);
                        int index = 0;
                        foreach (Match item in rgx.Matches(str1[i]))
                        {
                            string parametter = item.Value;
                            parametter = parametter.Replace(@"<li class=""jss210"">", "");
                            parametter = parametter.Replace(@"</li>", "");
                            parametter = parametter.Replace(@"<!-- -->", "");

                            if (index == 0)
                            {
                                products1[i].CountOfRooms = parametter;
                                index++;
                            }
                            else
                            {
                                products1[i].Metrazh = parametter;

                            }



                        }
                    }
                        patern = @"""updateTime"":""(.*?)"",""real";
                        rgx = new Regex(patern);
                        List<string> updates1 = new List<string>();
                        foreach (Match item in rgx.Matches(htmlCode))
                        {
                            updates1.Add(item.Value);
                        }
                        for (int l = 0; l < updates1.Count; l++)
                        {
                            updates1[l] = updates1[l].Replace(@"""updateTime"":""", "");
                            updates1[l] = updates1[l].Replace(@",""real", "");
                            products1[l].Updated = updates1[l];
                            //Console.WriteLine(updates[i]);
                        }
                        patern = @",""addTime"":""(.*?)"",""";
                        rgx = new Regex(patern);
                        List<string> addd1 = new List<string>();
                        foreach (Match item in rgx.Matches(htmlCode))
                        {
                            addd1.Add(item.Value);
                        }
                        for (int l = 0; l < addd1.Count; l++)
                        {
                            addd1[l] = addd1[l].Replace(@",""addTime"":""", "");
                            addd1[l] = addd1[l].Replace(@""",""", "");
                            products1[l].Created = addd1[l];
                            //Console.WriteLine(addd[i]);
                        }

                        products.AddRange(products1);
                    #endregion
                }
            }
            //MessageBox.Show(products.Count.ToString());
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                var excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];
                excelWorksheet.Cells[1, 1].Value = "Adress";
                excelWorksheet.Cells[1, 2].Value = "Price";
                excelWorksheet.Cells[1, 3].Value = "CountOfRooms";
                excelWorksheet.Cells[1, 4].Value = "Metrazh";
                excelWorksheet.Cells[1, 5].Value = "Link";
                excelWorksheet.Cells[1, 6].Value = "Created";
                excelWorksheet.Cells[1, 7].Value = "Updated";
                for (int i = 0; i < products.Count; i++)
                {
                    excelWorksheet.Cells[i+2, 1].Value = products[i].Adress;
                    excelWorksheet.Cells[i+2 ,2].Value = products[i].Price;
                    excelWorksheet.Cells[i + 2, 3].Value = products[i].CountOfRooms;
                    excelWorksheet.Cells[i + 2, 4].Value = products[i].Metrazh;
                    excelWorksheet.Cells[i + 2, 5].Value = products[i].Link;
                    excelWorksheet.Cells[i + 2, 6].Value = products[i].Created;
                    excelWorksheet.Cells[i + 2, 7].Value = products[i].Updated;
                   // MessageBox.Show(products[i].Adress);
                }
               
                FileInfo excelFile = new FileInfo(path+@"\"+filename+".xlsx");
                excel.SaveAs(excelFile);
                
            }
            MessageBox.Show("End parse");
            //string lastpath = path + @"\" + filename + ".xlsx";
            //Console.WriteLine(lastpath);
           
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
