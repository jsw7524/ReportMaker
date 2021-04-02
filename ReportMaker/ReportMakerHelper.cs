using OfficeOpenXml;
using System.IO;

namespace JswTools
{
    public class ReportMakerHelper
    {
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
    }




}