using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Xml.Linq;
using ODSFiles.Style;

namespace ODSFiles
{
    public class Row
    {
        internal XElement RowElement { get; set; }
        internal int Index { get; set; }
        internal string SheetName { get; set; }

        public void AddRowAfter()
        {
            RowElement.AddAfterSelf(new XElement(Globals.table + "table-row", new XAttribute(Globals.table + "style-name", "ro1")));
        }

        public void AddRowBefore()
        {
            RowElement.AddBeforeSelf(new XElement(Globals.table + "table-row", new XAttribute(Globals.table + "style-name", "ro1")));
        }

        public void DeleteRow()
        {
            RowElement.Remove();
        }

        public void Group()
        {
            if (!RowElement.ElementsBeforeSelf().Any())
            {
                RowElement.AddBeforeSelf(new XElement(Globals.table + "table-row-group", new XElement(RowElement)));
                RowElement.Remove();
            }
            else if (RowElement.ElementsBeforeSelf().Last().Name == Globals.table + "table-row-group")
            {
                RowElement.ElementsBeforeSelf().Last().Add(new XElement(RowElement));
                RowElement.Remove();
            }
            else
            {
                RowElement.AddBeforeSelf(new XElement(Globals.table + "table-row-group", new XElement(RowElement)));
                RowElement.Remove();
            }
        }

        //Не работает в Shell Разобраться
        public void LoadStyle(string cellType, CellStyle cellStyle)
        {
            IEnumerable<XElement> cellElements = RowElement.Elements(Globals.table + "table-cell");

            cellStyle.CreatElement(cellType);

            foreach (var cellElement in cellElements)
            {
                if (cellElement.Attribute(Globals.office + "value-type")?.Value == cellType)
                {
                    cellElement.SetAttributeValue(Globals.table + "style-name", $"ce{cellStyle.HashCodeStyle}");
                }
            }
        }

        public void LoadStyle(CellStyle cellStyle)
        {
            IEnumerable<XElement> cellElements = RowElement.Elements(Globals.table + "table-cell");

            foreach (var cellElement in cellElements)
            {
                cellStyle.CreatElement(cellElement.Attribute(Globals.office + "value-type")?.Value ?? "");
                cellElement.SetAttributeValue(Globals.table + "style-name", $"ce{cellStyle.HashCodeStyle}");
            }
        }

        public void LoadStyle(RowStyle rowStyle)
        {
            rowStyle.CreatElement();

            RowElement.SetAttributeValue(Globals.table + "style-name", $"ro{rowStyle.HashCodeStyle}");
        }

        public void FreezeRow()
        {
            IEnumerable<XElement> root = Globals.settings.Element(Globals.office + "document-settings")
                .Element(Globals.office + "settings").Elements(Globals.config + "config-item-set")
                .Where(obj => obj.Attribute(Globals.config + "name").Value == "ooo:view-settings").FirstOrDefault()
                .Element(Globals.config + "config-item-map-indexed").Element(Globals.config + "config-item-map-entry")
                .Element(Globals.config + "config-item-map-named").Elements(Globals.config + "config-item-map-entry");

            foreach (var elem in root)
            {
                if (elem.Attribute(Globals.config + "name").Value == $"{SheetName}")
                {
                    elem.Add(
                        new XElement(Globals.config + "config-item", "2",
                            new XAttribute(Globals.config + "name", "VerticalSplitMode"),
                            new XAttribute(Globals.config + "type", "short")),
                        new XElement(Globals.config + "config-item", $"{Index}",
                            new XAttribute(Globals.config + "name", "VerticalSplitPosition"),
                            new XAttribute(Globals.config + "type", "int")),
                        new XElement(Globals.config + "config-item", "3",
                            new XAttribute(Globals.config + "name", "ActiveSplitRange"),
                            new XAttribute(Globals.config + "type", "short")),
                        new XElement(Globals.config + "config-item", $"{Index}",
                            new XAttribute(Globals.config + "name", "PositionBottom"),
                            new XAttribute(Globals.config + "type", "int")));
                }
            }
        }
    }
}