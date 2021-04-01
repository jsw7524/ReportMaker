using System;
using System.Collections;
using System.Collections.Generic;

namespace JswTools
{
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

        public void Process(IEnumerable headData, IEnumerable bodyData, IEnumerable footData)
        {
            MakeHead(headData);
            MakeBody(bodyData, 0);
            MakeFoot(footData);
        }

        public void Process(IEnumerable data)
        {
            Process(data, data, data);
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
}