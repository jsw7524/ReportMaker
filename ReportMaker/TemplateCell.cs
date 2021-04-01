using System;

namespace JswTools
{
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
}