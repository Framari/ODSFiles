using System;

namespace ODSFiles
{
    public static class ConverterExcelCells
    {
        private const string ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static int ToNumeric(string index)
        {
            string first = string.Empty;

            CharEnumerator ce = index.GetEnumerator();
            while (ce.MoveNext())
                if (char.IsLetter(ce.Current))
                    first += ce.Current;

            int i = 0;
            ce = first.GetEnumerator();
            while (ce.MoveNext())
                i = (26 * i) + ALPHABET.IndexOf(ce.Current) + 1;

            return i;
        }

        public static int[] ToNumericCoordinates(string coordinates)
        {
            int[] coords = new int[2];
            string first = string.Empty;
            string second = string.Empty;

            CharEnumerator ce = coordinates.GetEnumerator();
            while (ce.MoveNext())
                if (char.IsLetter(ce.Current))
                    first += ce.Current;
                else
                    second += ce.Current;

            int i = 0;
            ce = first.GetEnumerator();
            while (ce.MoveNext())
                i = (26 * i) + ALPHABET.IndexOf(ce.Current) + 1;

            string str = i.ToString();

            coords[0] = Convert.ToInt32(str);
            coords[1] = Convert.ToInt32(second);

            return coords;
        }

        public static int[] RangeToNumericCoordinates(string coordinates)
        {
            int[] coords = new int[4];

            string[] coordinate = coordinates.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);

            coords[0] = ToNumericCoordinates(coordinate[0])[0];
            coords[1] = ToNumericCoordinates(coordinate[0])[1];
            coords[2] = ToNumericCoordinates(coordinate[1])[0];
            coords[3] = ToNumericCoordinates(coordinate[1])[1];

            return coords;
        }
    }
}