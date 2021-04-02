using System.IO;
using OfficeOpenXml;

namespace JswTools
{
    public static class DefaultStyleNames
    {
        public readonly static string DefaultStyles_無 = "B2";
        public readonly static string DefaultStyles_全框線 = "B4";
        public readonly static string DefaultStyles_下框線 = "B6";
    }
    public class StylesManager
    {
        private ExcelWorksheet _defaultStyleSheet;
        public StylesManager()
        {
            _defaultStyleSheet = LoadDefaultStyles();
        }
        public StylesManager(ExcelWorksheet sheet)
        {
            _defaultStyleSheet = sheet;
        }
        public ExcelWorksheet LoadDefaultStyles()
        {
            using (FileStream fileStream = new FileStream("DefaultStyles.xlsx", FileMode.Open))
            {
                MemoryStream memoryStream = new MemoryStream();
                ExcelPackage xls = new ExcelPackage(memoryStream, fileStream);
                return xls.Workbook.Worksheets["Styles"];
            }
        }

        public void ApplyStyle(string styleName, ExcelRange target)
        {
            ApplyStyle(_defaultStyleSheet, styleName, target);
        }

        public void ApplyStyle(ExcelWorksheet sheet, string indexer, ExcelRange target)
        {
            sheet.Cells[indexer].Copy(target);
        }

    }
}