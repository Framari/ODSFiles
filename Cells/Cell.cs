using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml.Linq;
using System.IO;
using System.IO.Compression;
using System.Linq;
using ODSFiles.Style;

namespace ODSFiles
{
    public class Cell
    {
        internal XElement CellElement { get; set; }

        public object Value
        {
            get
            {
                if (CellElement.Attribute(Globals.office + "value-type")?.Value == "float")
                {
                    return Convert.ToDouble(CellElement.Attribute(Globals.office + "value").Value);
                }
                if (CellElement.Attribute(Globals.office + "value-type")?.Value == "string")
                {
                    return CellElement.Element(Globals.text + "p").Value;
                }
                if (CellElement.Attribute(Globals.office + "value-type")?.Value == "currency")
                {
                    return Convert.ToDecimal(CellElement.Attribute(Globals.office + "value").Value);
                }
                if (CellElement.Attribute(Globals.office + "value-type")?.Value == "date")
                {
                    return Convert.ToDateTime(CellElement.Attribute(Globals.office + "value").Value);
                }

                return null;
            }

            set
            {
                if (value == null)
                {
                    CellElement.SetAttributeValue(Globals.office + "value-type", "string");
                    CellElement.Add(new XElement(Globals.text + "p", $""));
                }
                if (value.GetType() == typeof(int) || value.GetType() == typeof(float) || value.GetType() == typeof(double))
                {
                    CellElement.SetAttributeValue(Globals.office + "value-type", "float");
                    CellElement.SetAttributeValue(Globals.office + "value", $"{value}");
                }
                if (value.GetType() == typeof(string) || value.GetType() == typeof(char))
                {
                    CellElement.SetAttributeValue(Globals.office + "value-type", "string");
                    CellElement.Add(new XElement(Globals.text + "p", $"{value}"));
                }
                if (value.GetType() == typeof(decimal))
                {
                    CellElement.SetAttributeValue(Globals.office + "value-type", "currency");
                    CellElement.SetAttributeValue(Globals.office + "value", $"{value}");
                }
                if (value.GetType() == typeof(DateTime))
                {
                    CellElement.SetAttributeValue(Globals.office + "value-type", "date");
                    CellElement.SetAttributeValue(Globals.office + "date-value", $"{DateTime.Parse(value.ToString()).ToShortDateString()}");
                }
            }
        }

        public string Formula
        {
            set
            {
                CellElement.SetAttributeValue(Globals.office + "value-type", "float");
                CellElement.SetAttributeValue(Globals.office + "value-type", "currency");
                CellElement.SetAttributeValue(Globals.table + "formula", $"of:={value}");
            }
        }

        public void LoadStyle(CellStyle cellStyle)
        {
            cellStyle.CreatElement(CellElement.Attribute(Globals.office + "value-type")?.Value ?? "");

            CellElement.SetAttributeValue(Globals.table + "style-name", $"ce{cellStyle.HashCodeStyle}");
        }
    }
}