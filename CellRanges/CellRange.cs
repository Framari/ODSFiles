using System.Linq;
using System.Runtime.InteropServices;
using ODSFiles.Style;

namespace ODSFiles
{
    public class CellRange
    {
        private Tables table = new Tables(); 

        internal int Index1From { get; set; }
        internal int Index2From { get; set; }
        internal int Index1To { get; set; }
        internal int Index2To { get; set; }

        internal Cells Cells { get; set; }


        public void Merge()
        {
            int columns = Index2To - Index2From == 0 ? 1 : Index2To - Index2From + 1;
            int rows = Index1To - Index1From == 0 ? 1 : Index1To - Index1From + 1;

            Cells[Index1From, Index2From].CellElement.SetAttributeValue(Globals.table + "number-columns-spanned", $"{columns}");
            Cells[Index1From, Index2From].CellElement.SetAttributeValue(Globals.table + "number-rows-spanned", $"{rows}");
        }

        public Tables Table
        {
            get
            {
                table.Index1From = Index1From;
                table.Index2From = Index2From;
                table.Cells = Cells;
                return table;
            }
        }

        public object Value
        {
            set
            {
                if(Cells[Index1From,Index2From].CellElement.Attributes(Globals.table + "number-columns-spanned").Any() ||
                   Cells[Index1From, Index2From].CellElement.Attributes(Globals.table + "number-rows-spanned").Any())
                {
                    Cells[Index1From, Index2From].Value = value;
                }
            }
        }

        public void LoadStyle(CellStyle cellStyle)
        {
            for (int i = Index1From; i <= Index1To; i++)
            {
                for (int j = Index2From; j <= Index2To; j++)
                {
                    Cells[i, j].LoadStyle(cellStyle);
                }
            }
        }
    }
}