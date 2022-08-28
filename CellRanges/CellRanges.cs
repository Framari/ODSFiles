using System;
using System.Xml.Linq;

namespace ODSFiles
{
    public class CellRanges
    {
        private Cells Cells { get; }

        public CellRanges(Cells cells)
        {
            Cells = cells;
        }

        public CellRange this[string index]
        {
            get
            {
                return new CellRange()
                { 
                    Index1From = ConverterExcelCells.RangeToNumericCoordinates(index)[1],
                    Index1To = ConverterExcelCells.RangeToNumericCoordinates(index)[3],
                    Index2From = ConverterExcelCells.RangeToNumericCoordinates(index)[0],
                    Index2To = ConverterExcelCells.RangeToNumericCoordinates(index)[2],
                    Cells = Cells,

                };
            }
        }

        public CellRange this[int index1From, int index2From, int index1To, int index2To]
        {
            get
            {
                if (index1From <= 0 || index2From <= 0 || index1To <= 0 || index2To <= 0) throw new Exception("Ни один из индексов не может быть меньше 1");
               
                return new CellRange()
                {
                    Index1From = index1From,
                    Index1To = index1To,
                    Index2From = index2From,
                    Index2To = index2To,
                    Cells = Cells
                };
            }
        }
    }
}