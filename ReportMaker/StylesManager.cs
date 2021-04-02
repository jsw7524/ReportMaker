using System.IO;
using OfficeOpenXml;

namespace JswTools
{

    public class StylesManager
    {
        private ExcelWorksheet _defaultStyleSheet;

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