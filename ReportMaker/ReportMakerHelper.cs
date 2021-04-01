namespace JswTools
{
    public class ReportMakerHelper
    {
        public string ToROCYearMMDD(int yyyy, int mm, int dd)
        {
            return (yyyy - 1911).ToString() + "." + mm.ToString() + "." + dd.ToString();
        }
    }
}