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

        private StylesManager _stylesManager;

        public ReportMaker(StylesManager stylesManager = null, string deliminator1 = "%", string deliminator2 = "#")
        {
            _stylesManager = stylesManager;
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
                        foreach (var c in ws.Cells)
                        {
                            try
                            {
                                if (c.Value is null)
                                {
                                    continue;
                                }
                                Match match = tagFinder.Match(c.Value as string);
                                if (false == match.Success)
                                {
                                    continue;
                                }
                                string tag = match.Groups["tag"].Value;
                                for (int row = 0; row < data[tag].Count; row++)
                                {
                                    TemplateRow rowData = data[tag][row];
                                    if (false == string.IsNullOrEmpty(rowData.RowStyle))
                                    {
                                        _stylesManager?.ApplyStyle(rowData.RowStyle, ws.Cells[c.Start.Row + row, c.Start.Column, c.Start.Row + row, c.Start.Column + rowData.RowContent.Count]);
                                    }
                                    for (int col = 0; col < rowData.RowContent.Count; col++)
                                    {
                                        if (null != rowData.RowContent[col].DoSomething)
                                        {
                                            rowData.RowContent[col].DoSomething(rowData.RowContent[col]);
                                        }
                                        if (false == string.IsNullOrEmpty(rowData.RowContent[col].cellStyle))
                                        {
                                            _stylesManager?.ApplyStyle(rowData.RowContent[col].cellStyle, ws.Cells[row + c.Start.Row, col + c.Start.Column]);
                                        }
                                        ws.Cells[row + c.Start.Row, col + c.Start.Column].Value = rowData.RowContent[col].content;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                c.Value = c.Value + ":" + ex;
                            }
                        }
                    }
                    xls.Save();
                }
            }
            return memoryStream;
        }
    }
}