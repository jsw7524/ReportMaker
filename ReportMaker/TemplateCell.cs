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
        public TemplateCell() : this("")
        {
            return;
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
}