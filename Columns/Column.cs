using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ODSFiles.Style;

namespace ODSFiles
{
    public class Column
    {
        internal IEnumerable<XElement> CellElements { get; set; }
        internal XElement ColumnElement { get; set; }

        internal int Index { get; set; }
        internal string SheetName { get; set; }

        public void AddColumnBefore()
        {
            foreach (var cell in CellElements)
            {
                cell.AddBeforeSelf(new XElement(Globals.table + "table-cell"));
            }
        }

        public void AddColumnAfter()
        {
            foreach (var cell in CellElements)
            {
                cell.AddAfterSelf(new XElement(Globals.table + "table-cell"));
            }
        }

        public void DeleteColumn()
        {
            foreach (var cell in CellElements)
            {
                cell.Remove();
            }
        }

        //Не работает в Shell разобраться
        public void LoadStyle(CellStyle cellStyle)
        {
            foreach (var cellElement in CellElements)
            {
                cellStyle.CreatElement(cellElement.Attribute(Globals.office + "value-type")?.Value ?? "");
                cellElement.SetAttributeValue(Globals.table + "style-name", $"ce{cellStyle.HashCodeStyle}");
            }
        }

        public void LoadStyle(string cellType, CellStyle cellStyle)
        {
            cellStyle.CreatElement(cellType);

            foreach (var cellElement in CellElements)
            {
                if (cellElement.Attribute(Globals.office + "value-type")?.Value == cellType)
                {
                    cellElement.SetAttributeValue(Globals.table + "style-name", $"ce{cellStyle.HashCodeStyle}");
                }
            }
        }

        public void LoadStyle(ColumnStyle columnStyle)
        {
            columnStyle.CreatElement();
            ColumnElement.SetAttributeValue(Globals.table + "style-name", $"co{columnStyle.HashCodeStyle}");
        }

        public void FreezeColumn()
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
                            new XAttribute(Globals.config + "name", "HorizontalSplitMode"),
                            new XAttribute(Globals.config + "type", "short")),
                        new XElement(Globals.config + "config-item", $"{Index}",
                            new XAttribute(Globals.config + "name", "HorizontalSplitPosition"),
                            new XAttribute(Globals.config + "type", "int")),
                        new XElement(Globals.config + "config-item", "3",
                            new XAttribute(Globals.config + "name", "ActiveSplitRange"),
                            new XAttribute(Globals.config + "type", "short")),
                        new XElement(Globals.config + "config-item", $"{Index}",
                            new XAttribute(Globals.config + "name", "PositionRight"),
                            new XAttribute(Globals.config + "type", "int")));
                }
            }
        }
    }
}