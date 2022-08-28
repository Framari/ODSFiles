using System.Collections.Generic;
using System.Threading;
using System.Xml.Linq;

namespace ODSFiles.Style
{
    public class ColumnStyle
    {
        internal int HashCodeStyle { get; set; }

        internal string StyleName { get; set; }

        public string Width { internal get; set; }

        internal void CreatElement()
        {
            StyleName = $"{Width ?? "2.258cm"}";

            HashCodeStyle = StyleName.GetHashCode();

            bool exist = false;

            IEnumerable<XElement> elements = Globals.content.Element(Globals.office + "document-content")
                .Element(Globals.office + "automatic-styles").Elements(Globals.style + "style");

            foreach (var elem in elements)
            {
                if (elem.Attribute(Globals.style + "name").Value == $"co{HashCodeStyle}")
                {
                    exist = true;
                    break;
                }
            }

            if (!exist)
            {
                XElement columnStyleElement = new XElement(Globals.style + "style",
                    new XAttribute(Globals.style + "name", $"co{HashCodeStyle}"),
                    new XAttribute(Globals.style + "family", "table-column")
                );

                columnStyleElement.Add(new XElement(Globals.style + "table-column-properties",
                    new XAttribute(Globals.fo + "break-before", "auto"),
                    new XAttribute(Globals.style + "column-width", $"{Width ?? "2.258cm"}")
                   
                ));

                Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "automatic-styles").Add(columnStyleElement);
            }
        }
    }
}