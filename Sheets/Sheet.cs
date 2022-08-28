using System.Collections;
using System.Collections.Generic;
using System.Xml.Linq;

namespace ODSFiles
{
    public class Sheet
    {
        private XElement SheetElement { get; }
        private Rows Rows { get; }
        private Cells Cells { get; }
        private Columns Columns { get; }
        private CellRanges CellRanges { get; }

        internal string SheetName { get; set; }

        public Sheet(string name, XElement sheetElement)
        {
            SheetName = name;
            SheetElement = sheetElement;
            Rows = new Rows(SheetElement, SheetName);
            Cells = new Cells(Rows);
            CellRanges = new CellRanges(Cells);
            Columns = new Columns(SheetElement, Cells, SheetName);
        }

        public Cells Cell
        {
            get { return Cells; }
        }

        public Rows Row
        {
            get { return Rows; }
        }

        public Columns Column
        {
            get { return Columns; }
        }

        public CellRanges CellRange
        {
            get { return CellRanges; }
        }
    }
}