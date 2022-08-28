using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ODSFiles
{
    public class Sheets
    {
        private List<Sheet> _sheetList = new List<Sheet>();

        public Sheet this[string name]
        {
            get { return GetSheet(name); }
        }

        private Sheet GetSheet(string sheetName)
        {
            IEnumerable<XElement> sheets = Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "body").Element(Globals.office + "spreadsheet").Elements(Globals.table + "table");

            foreach (XElement element in sheets)
            {
                if (element.FirstAttribute.Value == $"{sheetName}")
                {
                    _sheetList.Add(new Sheet(sheetName, element));
                    return _sheetList.Find(s => s.SheetName == sheetName);
                }
            }

            throw new Exception("Такая страница не существует");
        }

        public void CreateSheet(string sheetName)
        {
            XElement root = Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "body").Element(Globals.office + "spreadsheet");

            root.Add(new XElement(Globals.table + "table",
                new XAttribute(Globals.table + "name", $"{sheetName}")));

            XElement configSettings = Globals.settings.Element(Globals.office + "document-settings")
                .Element(Globals.office + "settings").Elements(Globals.config + "config-item-set")
                .Where(obj => obj.Attribute(Globals.config + "name").Value == "ooo:configuration-settings")
                .FirstOrDefault()
                .Element(Globals.config + "config-item-map-named");

            configSettings.Add(new XElement(Globals.config + "config-item-map-entry",
                new XAttribute(Globals.config + "name", $"{sheetName}"),
                new XElement(Globals.config + "config-item", $"{sheetName}",
                    new XAttribute(Globals.config + "name", "CodeName"),
                    new XAttribute(Globals.config + "type", "string"))));

            XElement viewSettings = Globals.settings.Element(Globals.office + "document-settings")
                .Element(Globals.office + "settings").Elements(Globals.config + "config-item-set")
                .Where(obj => obj.Attribute(Globals.config + "name").Value == "ooo:view-settings").FirstOrDefault()
                .Element(Globals.config + "config-item-map-indexed").Element(Globals.config + "config-item-map-entry")
                .Element(Globals.config + "config-item-map-named");

            viewSettings.Add(new XElement(Globals.config + "config-item-map-entry",
                new XAttribute(Globals.config + "name", $"{sheetName}")));

        }
    }
}