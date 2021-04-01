using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using OfficeOpenXml;
using System.Data;

namespace JswTools
{
    public class ReportMaker
    {
        private string _deliminator1;
        private string _deliminator2;
        public ReportMaker(string deliminator1 = "%", string deliminator2 = "#")
        {
            _deliminator1 = deliminator1;
            _deliminator2 = deliminator2;
        }

        public MemoryStream FillDataInTemplate(string templatefilename, IDictionary<string, List<TemplateRow>> data)
        {
            Regex tagFinder = new Regex("^" + _deliminator1 + "(?<tag>.+)" + _deliminator2 + "$");
            MemoryStream memoryStream = new MemoryStream();
            using (FileStream templateFileStream = new FileStream(templatefilename, FileMode.Open))
            {
                using (ExcelPackage xls = new ExcelPackage(memoryStream, templateFileStream))
                {
                    var sheets = xls.Workbook.Worksheets;
                    foreach (var ws in sheets)
                    {
                        if ("StylesJSW" == ws.Name)
                            continue;
                        foreach (var c in ws.Cells)
                        {
                            try
                            {
                                var match = tagFinder.Match(c.Value as string);
                                if (!match.Success)
                                {
                                    continue;
                                }
                                string tag = match.Groups["tag"].Value;
                                for (int row = 0; row < data[tag].Count; row++)
                                {
                                    var rowData = data[tag][row];
                                    for (int col = 0; col < rowData.RowContent.Count; col++)
                                    {
                                        ws.Cells[row + c.Start.Row, col + c.Start.Column].Value = rowData.RowContent[col].content;
                                        if ("" != rowData.RowContent[col].StyleInfo)
                                            sheets["StylesJSW"].Cells[rowData.RowContent[col].StyleInfo].Copy(ws.Cells[row + c.Start.Row, col + c.Start.Column]);
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                c.Value = c.Value + " not found";
                            }
                        }
                    }
                    if (null != sheets["StylesJSW"])
                    {
                        sheets.Delete("StylesJSW");
                    }
                    xls.Save();
                }
            }
            return memoryStream;
        }
    }
}