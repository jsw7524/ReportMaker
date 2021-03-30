using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.Web.Configuration;

namespace JswTools
{
    public class ToolsJsw
    {
        public string ToYYYMMDD(int yyyy, int mm, int dd)
        {
            return (yyyy - 1911).ToString() + "." + mm.ToString() + "." + dd.ToString();
        }
    }

    /////////////////////////////////////////////////
    public class Template
    {
        Dictionary<string, ExcelRange> stylesDict = new Dictionary<string, ExcelRange>();

        public MemoryStream FillReport(string filename, string templatefilename, Dictionary<string, List<TemplateRow>> data)
        {
            return FillReport(filename, templatefilename, data, new string[] { "%", "#" });
        }

        public MemoryStream FillReport(string filename, string templatefilename, Dictionary<string, List<TemplateRow>> data, string[] deliminator)
        {
            var file = new MemoryStream();
            using (var temp = new FileStream(templatefilename, FileMode.Open))
            {
                using (var xls = new ExcelPackage(file, temp))
                {
                    foreach (var ws in xls.Workbook.Worksheets)
                    {
                        if ("StylesJSW" == ws.Name)
                            continue;
                        foreach (var c in ws.Cells)
                        {
                            var dataLabel = "" + c.Value;
                            if (dataLabel.StartsWith(deliminator[0]) == false || dataLabel.EndsWith(deliminator[1]) == false)
                            {
                                continue;
                            }
                            dataLabel = dataLabel.Replace(deliminator[0], "").Replace(deliminator[1], "");
                            try
                            {
                                for (int row = 0; row < data[dataLabel].Count; row++)
                                {
                                    var rowData = data[dataLabel][row];
                                    for (int col = 0; col < rowData.RowContent.Count ; col++)
                                    {
                                        if ("" == rowData.RowContent[col].info)
                                            continue;
                                        xls.Workbook.Worksheets["StylesJSW"].Cells[rowData.RowContent[col].info].Copy(ws.Cells[row+c.Start.Row, col+c.Start.Column]);
                                    }
                                }
                                var reportData = ((data[dataLabel]).Select(a => a.RowContent).Select(b => b.Select(z => z.content as object).ToArray()).ToArray());
                                c.LoadFromArrays(reportData);
                            }
                            catch (Exception ex)
                            {
                                c.Value = dataLabel + " not found";
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
            return file;

        }
    }
    /////////////////////////////////////////////////

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
            info = t;
        }

        public object content;
        public string info;
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

    public class DataProvider
    {
        public DataTable GetData(SqlCommand cmd)
        {
            cmd.Connection = new SqlConnection(WebConfigurationManager.ConnectionStrings["福利會資料庫ConnectionString1"].ConnectionString.ToString());
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }
    }
}