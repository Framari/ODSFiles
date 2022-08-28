using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ODSFiles
{
    public class Cells
    {
        private Rows Rows { get; }

        public Cells(Rows rows)
        {
            Rows = rows;
        }

        public Cell this[string index]
        {
            get
            {
                return new Cell()
                {
                    CellElement = GetCellElement(
                        Rows.GetRowElement(
                            ConverterExcelCells.ToNumericCoordinates(index)[1]),
                            ConverterExcelCells.ToNumericCoordinates(index)[0])
                };
            }
        }

        public Cell this[int index1, int index2]
        {
            get
            {
                if (index1 <= 0 || index2 <= 0) throw new Exception("Индекс колонки или столбца не может быть меньше 1");
                
                return new Cell()
                {
                    CellElement = GetCellElement(Rows.GetRowElement(index1), index2)
                }; 
            }
        }

        private void SetCellElement(XElement row , int cellIndex)
        {
            IEnumerable<XElement> elements = row.Elements(Globals.table + "table-cell");

            if (cellIndex == 0)
            {
                row.Add(new XElement(Globals.table + "table-cell"));
            }
            else
            {
                elements.ElementAt(cellIndex).AddAfterSelf(new XElement(Globals.table + "table-cell"));
            }
        }

        internal XElement GetCellElement(XElement row ,int cellIndex)
        {
            cellIndex--;

            IEnumerable<XElement> elements = row.Elements(Globals.table + "table-cell");

            int length = elements.Count() - 1;

            if (length < cellIndex)
            {
                while (length < cellIndex)
                {
                    SetCellElement(row , length == -1 ? 0 : length);
                    length++;
                }

                return elements.ElementAt(cellIndex);
            }
            else
            {
                return elements.ElementAt(cellIndex);
            }
        }
    }
}