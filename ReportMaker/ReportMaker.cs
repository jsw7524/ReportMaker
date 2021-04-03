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
        private ReportMakerHelper _reportMakerHelper;

        public ReportMaker(StylesManager stylesManager = null)
        {
            _stylesManager = stylesManager;
            _reportMakerHelper = new ReportMakerHelper();
        }

        public MemoryStream FillDataInTemplate(string templatefilename, IDictionary<string, List<TemplateRow>> data)
        {
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
                                string tag = _reportMakerHelper.FindCellTag(c);
                                if (null == tag || false==data.ContainsKey(tag))
                                {
                                    continue;
                                }
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
                                throw ex;
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