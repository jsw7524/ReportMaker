using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace JswTools
{

    public class StylesManager
    {
        private ExcelWorksheet _defaultStyleSheet;
        private Dictionary<string, string> _stylesInfo;
        private ReportMakerHelper _reportMakerHelper;

        public StylesManager(ExcelWorksheet sheet)
        {
            _defaultStyleSheet = sheet;
            _stylesInfo = new Dictionary<string, string>();
            _reportMakerHelper = new ReportMakerHelper();
            LoadStylesInfo();
        }
        public void LoadStylesInfo()
        {
            if (null == _defaultStyleSheet)
            {
                return;
            }
            foreach (var cell in _defaultStyleSheet.Cells)
            {
                string tag = _reportMakerHelper.FindCellTag(cell);
                if (null == tag)
                {
                    continue;
                }
                string styleInfo = _reportMakerHelper.FindCellStyle(cell);
                if (null != styleInfo)
                {
                    _stylesInfo.Add(tag, styleInfo);
                }
            }
            return;
        }

        public void ApplyStyle(string styleName, ExcelRange target)
        {
            if (_stylesInfo.ContainsKey(styleName))
            {
                ApplyStyle(_defaultStyleSheet, _stylesInfo[styleName], target);
                return;
            }
            ApplyStyle(_defaultStyleSheet, styleName, target);
        }

        public void ApplyStyle(ExcelWorksheet sheet, string indexer, ExcelRange target)
        {
            sheet.Cells[indexer].Copy(target);
        }

    }
}