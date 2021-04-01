using System.Collections;
using System.Collections.Generic;

namespace JswTools
{
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
}