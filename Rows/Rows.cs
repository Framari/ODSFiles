using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ODSFiles
{
    public class Rows
    {
        private XElement SheetElement { get;}

        private string SheetName { get; set; }

        public Rows(XElement sheet, string sheetName)
        {
            SheetName = sheetName;
            SheetElement = sheet;
        }

        public Row this[int index]
        {
            get
            {
                if (index <= 0) throw new Exception("Индекс строки не может быть меньше 1");

                return new Row()
                {
                    RowElement = GetRowElement(index),
                    SheetName = SheetName,
                    Index = index
                };
            }
        }

        private void SetRowElement(int rowIndex)
        {
            IEnumerable<XElement> elements = SheetElement.Descendants(Globals.table + "table-row");

            if (rowIndex == 0)
            {
                SheetElement.Add(new XElement(Globals.table + "table-row", new XAttribute(Globals.table + "style-name", "ro1")));
            }
            else
            {
                elements.ElementAt(rowIndex).AddAfterSelf(new XElement(Globals.table + "table-row", new XAttribute(Globals.table + "style-name", "ro1")));
            }
        }

        internal XElement GetRowElement(int rowIndex)
        {
            rowIndex--;
            IEnumerable<XElement> elements = SheetElement.Descendants(Globals.table + "table-row");

            int length = elements.Count() - 1;

            if (length < rowIndex)
            {
                while (length != rowIndex)
                {
                    SetRowElement(length == -1 ? 0 : length);
                    length++;
                }

                return elements.ElementAt(rowIndex);
            }
            else
            {
                return elements.ElementAt(rowIndex);
            }
        }
    }
}