using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Supremes;
using OfficeOpenXml;
using System.IO;


namespace MySoup
{
    class Program
    {
        static void Main(string[] args)
        {
            File.Delete("Curtains.xlsx");
            File.Delete("Furnitures.xlsx");
            
            var curtains = new List<Data>();
            var furnitures = new List<Data>();

            for (int i = 1; i <= 7; i++)
            {
                var url = "https://trangvangvietnam.com/tagclass/30074810/rem-cua.html?page=" + i.ToString();
                Console.WriteLine(url);
                //var getInforTask = Task<List<Data>>.Run(() => GetInfo(url));
                var result = GetInfo(url).GetAwaiter().GetResult();
                curtains.AddRange(result);

            }
            Task.Run(() => WriteToExcel(curtains, "Curtains.xlsx"));
            Console.WriteLine("Write Curtains complete");
            for (int i = 1; i <= 27; i++)
            {
                var url = "https://trangvangvietnam.com/srch/n%E1%BB%99i_th%E1%BA%A5t.html?page=" + i.ToString();
                Console.WriteLine(url);
                var getInforTask = GetInfo(url).GetAwaiter().GetResult();
                furnitures.AddRange(getInforTask);
               
            }
            Task.Run(() => WriteToExcel(furnitures, "Furnitures.xlsx"));
            Console.WriteLine("Write Furniture complete");
        }

        public static async Task<string> GetTaxNumber(string url)
        {
            double flag;
            var result = string.Empty;
            var doc1 = Dcsoup.Parse(new Uri(url), 10000);
            foreach (var item in doc1.Select("div[class=hosocongty_text]"))
            {
                var taxString = item.Text;
                if (taxString.Length >= 7 && double.TryParse(taxString, out flag))
                {
                    result = taxString;
                    break;
                }
            }
            var t =  await Task.FromResult<string>(result);
            return t;
        }

        public static async Task<List<Data>> GetInfo(string url)
        {
            var data = new List<Data>();
            try
            {
                var doc1 = Dcsoup.Parse(new Uri(url), 10000);

                foreach (var item in doc1.Select("div[class=boxlistings]"))
                {
                    var address = item.Select("p[class=diachisection]").Last.Text;
                    if (address.Contains("TPHCM") || address.Contains("Đồng Nai") || address.Contains("Bình Dương") ||
                        address.Contains("Bình Phước") || address.Contains("Tây Ninh") || address.Contains("Long An"))
                    {
                        var newCurtain = new Data()
                        {
                            Name = item.Select("h2").Text,
                            DetailUrl = item.Select("a[class=buttonMoreDetails]").Attr("href"),
                            Address = address,
                            TaxNumber = ""
                        };
                        var result = await GetTaxNumber(newCurtain.DetailUrl);
                        if (!string.IsNullOrEmpty(result))
                        {
                            newCurtain.TaxNumber = result;
                            data.Add(newCurtain);
                        }

                    }
                }
            }
            catch(AggregateException ex)
            {
                foreach (var err in ex.InnerExceptions)
                {
                    Console.WriteLine(err.Message);
                }
                
            }
            return data;
        }
        public static void WriteToExcel(List<Data> data, string filename)
        {
           
            using (var package = new ExcelPackage(new FileInfo(filename)))
            {

                var ws = package.Workbook.Worksheets.Add("Sheet1");
                try
                {
                    ws.Cells["A1"].LoadFromCollection(data);
                    ws.Column(1).Width = 100;
                    ws.Column(2).Width = 100;
                    ws.Column(3).Width = 20;
                    ws.Column(4).Width = 100;
                    package.Save();
                }
                catch (AggregateException ex)
                {
                    foreach (var err in ex.InnerExceptions)
                    {
                        Console.WriteLine(err.Message);
                    }
                    
                }
            }
        }
    }
   
}
