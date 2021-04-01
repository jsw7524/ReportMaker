using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using OfficeOpenXml;
using System.Data;

namespace JswTools
{
    public class ToolsJsw
    {
        public string ToYYYMMDD(int yyyy, int mm, int dd)
        {
            return (yyyy - 1911).ToString() + "." + mm.ToString() + "." + dd.ToString();
        }
    }
    public class ReportMaker
    {
        private string _deliminator1 = "%";
        private string _deliminator2 = "#";
        public ReportMaker(string deliminator1= "%", string deliminator2= "#")
        {
            _deliminator1 = deliminator1;
            _deliminator2 = deliminator2;
        }

        public MemoryStream FillReport(string templatefilename, IDictionary<string, List<TemplateRow>> data)
        {
            Regex tagFinder = new Regex("^"+_deliminator1+"(?<tag>.*)"+_deliminator2+"$");
            MemoryStream memoryStream = new MemoryStream();
            using (FileStream templateFileStream = new FileStream(templatefilename, FileMode.Open))
            {
                using (ExcelPackage xls = new ExcelPackage(memoryStream, templateFileStream))
                {
                    foreach (var ws in xls.Workbook.Worksheets)
                    {
                        if ("StylesJSW" == ws.Name)
                            continue;
                        foreach (var c in ws.Cells)
                        {
                            var match = tagFinder.Match(c.Value as string);
                            if (!match.Success)
                            {
                                continue;
                            }
                            string tag = match.Groups["tag"].Value;
                            try
                            {
                                for (int row = 0; row < data[tag].Count; row++)
                                {
                                    var rowData = data[tag][row];
                                    for (int col = 0; col < rowData.RowContent.Count; col++)
                                    {
                                        if ("" == rowData.RowContent[col].StyleInfo)
                                            continue;
                                        xls.Workbook.Worksheets["StylesJSW"].Cells[rowData.RowContent[col].StyleInfo].Copy(ws.Cells[row + c.Start.Row, col + c.Start.Column]);
                                    }
                                }
                                var reportData = ((data[tag]).Select(a => a.RowContent).Select(b => b.Select(z => z.content as object).ToArray()).ToArray());
                                c.LoadFromArrays(reportData);
                            }
                            catch (Exception ex)
                            {
                                c.Value = tag + " not found";
                            }
                        }
                    }
                    if (null != xls.Workbook.Worksheets["StylesJSW"])
                    {
                        xls.Workbook.Worksheets.Delete("StylesJSW");
                    }
                    xls.Save();
                }
            }
            return memoryStream;
        }
    }
    public class TemplateCell
    {
        public static implicit operator TemplateCell(Decimal dec)
        {
            return new TemplateCell(dec);
        }
        public static implicit operator TemplateCell(string str)
        {
            return new TemplateCell(str);
        }
        public TemplateCell(object c) : this(c, "")
        {
            return;
        }
        public TemplateCell(object c, string t)
        {
            content = c;
            StyleInfo = t;
        }
        public object content;
        public string StyleInfo;
    }

    public class TemplateRow : IEnumerable
    {
        public List<TemplateCell> RowContent { set; get; }
        public string RowStyle { set; get; }
        public bool NewPage { set; get; }
        public TemplateRow()
        {
            RowContent = new List<TemplateCell>();
            RowStyle = "";
            NewPage = false;
        }
        public IEnumerator GetEnumerator()
        {
            return RowContent.GetEnumerator();
        }
    }

    public class DataProcessor
    {
        public List<TemplateRow> Result { get; set; }
        public Action<IEnumerable, int> BeforeGroup = (x, d) => { };
        public Action<IEnumerable, int> AfterGroup = (x, d) => { };
        public Action<object, int> RowDetail = (x, d) => { };
        public Action<IEnumerable> MakeHead = x => { };
        public Action<IEnumerable> MakeFoot = x => { };

        public DataProcessor()
        {
            Result = new List<TemplateRow>();
        }

        public void Process(IEnumerable data)
        {
            MakeHead(data);
            MakeBody(data, 0);
            MakeFoot(data);
        }

        public void MakeBody(IEnumerable data, int level)
        {
            foreach (var d in data)
            {
                if (d is IEnumerable)
                {
                    BeforeGroup(d as IEnumerable, level + 1);
                    MakeBody(d as IEnumerable, level + 1);
                    AfterGroup(d as IEnumerable, level + 1);
                }
                else
                {
                    RowDetail(d, level);
                }
            }
        }
    }
}