using Example;
using Example.Model;
using JswTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Example
{
    class AdventureWorksManager
    {
        AdventureWorksLTEntities db = new AdventureWorksLTEntities();
        public IEnumerable<Product> GetProductData()
        {
            return db.Product.AsEnumerable();
        }

        public void MakeReport1()
        {
            List<Product> Products = GetProductData().ToList();
            DataProcessor dp = new DataProcessor();
            dp.BeforeGroup = (x, d) =>
             {
                 if (1 == d)
                 {
                     var groupColor = x as IGrouping<string, Product>;
                     TemplateRow tmp = new TemplateRow();
                     TemplateCell colorTemplateCell = new TemplateCell() { content = "Color:", cellStyle = "B2" };
                     TemplateCell KeyTemplateCell = new TemplateCell() { content = groupColor.Key, cellStyle = "B2" };

                     tmp.RowContent.AddRange(new List<TemplateCell>() { colorTemplateCell, KeyTemplateCell });
                     dp.Result.Add(tmp);
                 }
             };

            dp.RowDetail = (x, d) =>
            {
                var p = x as Product;
                TemplateRow tmp = new TemplateRow();
                tmp.RowContent.AddRange(new List<TemplateCell>() { p.Name, p.Weight ?? 0m, p.Size });
                dp.Result.Add(tmp);
            };

            dp.Process(Products.GroupBy(a => a.Color));
            ReportMaker reportMaker = new ReportMaker();
            Dictionary<string, List<TemplateRow>> dict = new Dictionary<string, List<TemplateRow>>() { { "Products", dp.Result } };
            var image = reportMaker.FillDataInTemplate("Template1.xlsx", dict);
            File.WriteAllBytes("Test1.xlsx", image.ToArray());
        }


        public void MakeReport2()
        {
            List<Product> Products = GetProductData().ToList();
            DataProcessor dp = new DataProcessor();
            dp.BeforeGroup = (x, d) =>
            {
                if (1 == d)
                {
                    var groupColor = x as IGrouping<string, Product>;
                    TemplateRow tmp = new TemplateRow();
                    TemplateCell colorTemplateCell = new TemplateCell() { content = "Color:", cellStyle = "B2" };
                    tmp.RowContent.AddRange(new List<TemplateCell>() { colorTemplateCell, groupColor.Key, groupColor.AsEnumerable().Count() });
                    dp.Result.Add(tmp);
                }
            };
            dp.Process(Products.GroupBy(a => a.Color));
            ReportMaker reportMaker = new ReportMaker();
            Dictionary<string, List<TemplateRow>> dict = new Dictionary<string, List<TemplateRow>>() { { "Products", dp.Result } };
            var image = reportMaker.FillDataInTemplate("Template1.xlsx", dict);
            File.WriteAllBytes("Test2.xlsx", image.ToArray());
        }

        public void MakeReport3()
        {
            List<Product> Products = GetProductData().ToList();
            DataProcessor dp = new DataProcessor();
            dp.BeforeGroup = (x, d) =>
            {
                if (1 == d)
                {
                    var groupColor = x as IGrouping<string, Product>;
                    TemplateRow tmp = new TemplateRow();
                    tmp.RowStyle = "B6:D6";
                    tmp.RowContent.AddRange(new List<TemplateCell>() { "Color:", groupColor.Key, groupColor.AsEnumerable().Count() });
                    dp.Result.Add(tmp);
                }
            };

            dp.Process(Products.GroupBy(a => a.Color));

            ReportMakerHelper reportMakerhelper = new ReportMakerHelper();

            MemoryStream msStyleSheet = new MemoryStream();
            var styleExcel=reportMakerhelper.GetExcelInMemory("Template2.xlsx", msStyleSheet);
            var styleSheet= styleExcel.Workbook.Worksheets["Styles"];

            ReportMaker reportMaker = new ReportMaker(new StylesManager(styleSheet));
            Dictionary<string, List<TemplateRow>> dict = new Dictionary<string, List<TemplateRow>>() { { "Products", dp.Result } };
            var image = reportMaker.FillDataInTemplate("Template2.xlsx", dict);
            File.WriteAllBytes("Test3.xlsx", image.ToArray());
        }

    }


}




class Program
{
    static void Main(string[] args)
    {
        AdventureWorksManager awm = new AdventureWorksManager();
        //awm.MakeReport1();
        //awm.MakeReport2();
        awm.MakeReport3();
    }
}

