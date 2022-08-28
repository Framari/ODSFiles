using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ODSFiles.Style;

namespace ODSFiles
{
    public class Columns
    {
        private XElement SheetElement { get; }
        private string SheetName { get; set; }
        private Cells Cell { get; }

        public Columns(XElement sheet, Cells cells, string sheetName)
        {
            SheetName = sheetName;
            SheetElement = sheet;
            Cell = cells;
        }

        public Column this[string index]
        {
            get
            {
                return new Column()
                {
                    ColumnElement = GetColumnElement(ConverterExcelCells.ToNumeric(index)),
                    CellElements = GetCellElements(ConverterExcelCells.ToNumeric(index)),
                    Index = ConverterExcelCells.ToNumeric(index),
                    SheetName = SheetName
                };
            }
        }

        public Column this[int index]
        {
            get
            {
                if (index <= 0) throw new Exception("Индекс колонки не может быть меньше 1");
                
                return new Column()
                {
                    ColumnElement = GetColumnElement(index),
                    CellElements = GetCellElements(index),
                    Index = index,
                    SheetName = SheetName
                };
            }
        }

        private void SetColumnElement(IEnumerable<XElement> columnsElements, int columnIndex)
        {
            while (columnsElements.ElementAtOrDefault(columnIndex - 1) == null)
            {
                SheetElement.AddFirst(new XElement(Globals.table + "table-column", new XAttribute(Globals.table + "style-name", "co1"), new XAttribute(Globals.table + "default-cell-style-name", "ce1")));
            }
        }

        internal IEnumerable<XElement> GetCellElements(int columnIndex)
        {
            IEnumerable<XElement> rowsElements = SheetElement.Descendants(Globals.table + "table-row");

            IList<XElement> cellElements = new List<XElement>();

            for (int i = 0; i < rowsElements.Count(); i++)
            {
                cellElements.Add(Cell[i + 1, columnIndex].CellElement);
            }

            return cellElements.AsEnumerable();
        }

        internal XElement GetColumnElement(int columnIndex)
        {
            IEnumerable<XElement> columnsElements = SheetElement.Elements(Globals.table + "table-column");

            SetColumnElement(columnsElements, columnIndex);

            return columnsElements.ElementAt(columnIndex - 1);
        }
    }
}