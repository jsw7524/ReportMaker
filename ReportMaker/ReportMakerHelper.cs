using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System.IO;
using System.Text.RegularExpressions;

namespace JswTools
{
    public class ReportMakerHelper
    {

        Regex infoFinder = new Regex("^%(?<tag>[a-zA-Z][a-zA-Z1-9]*)(!(?<style>.+))?#$");

        public string ToROCYearMMDD(int yyyy, int mm, int dd)
        {
            return (yyyy - 1911).ToString() + "." + mm.ToString() + "." + dd.ToString();
        }

        public ExcelPackage GetExcelInMemory(string fileName, MemoryStream memoryStream)
        {
            using (FileStream templateFileStream = new FileStream(fileName, FileMode.Open))
            {
                ExcelPackage xlsx = new ExcelPackage(memoryStream, templateFileStream);
                return xlsx;
            }
        }

        public ExcelWorksheet GetSheet(string xlsxFile, string sheetName)
        {
            MemoryStream msStyleSheet = new MemoryStream();
            var styleExcel = GetExcelInMemory(xlsxFile, msStyleSheet);
            var styleSheet = styleExcel.Workbook.Worksheets[sheetName];
            return styleSheet;
        }


        public string FindCellTag(ExcelRangeBase c)
        {
            return FindCellInfo(c, "tag");
        }
        public string FindCellStyle(ExcelRangeBase c)
        {
            return FindCellInfo(c, "style");
        }
        public string FindCellInfo(ExcelRangeBase c, string info)
        {
            if (c.Value is null)
            {
                return null;
            }
            Match match = infoFinder.Match(c.Value as string);
            if (false == match.Success)
            {
                return null;
            }
            string style = match.Groups[info].Value;
            return style;
        }
    }




}