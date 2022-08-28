using System.Collections.Generic;
using System.Threading;
using System.Xml.Linq;

namespace ODSFiles.Style
{
    public class RowStyle
    {
        internal int HashCodeStyle { get; set; }

        internal string StyleName { get; set; }

        public string Height { internal get; set; }


        internal void CreatElement()
        {
            string optimalRowHeight = Height == null 
                    ? "true"
                    : "false";

            StyleName = $"{Height ?? "0.452cm"}{optimalRowHeight}";

            HashCodeStyle = StyleName.GetHashCode();

            bool exist = false;

            IEnumerable<XElement> elements = Globals.content.Element(Globals.office + "document-content")
                .Element(Globals.office + "automatic-styles").Elements(Globals.style + "style");

            foreach (var elem in elements)
            {
                if (elem.Attribute(Globals.style + "name").Value == $"ro{HashCodeStyle}")
                {
                    exist = true;
                    break;
                }
            }

            if (!exist)
            {
                 XElement rowStyleElement = new XElement(Globals.style + "style", 
                    new XAttribute(Globals.style + "name", $"ro{HashCodeStyle}"),
                    new XAttribute(Globals.style + "family", "table-row")
                    );

                rowStyleElement.Add(new XElement(Globals.style + "table-row-properties", 
                    new XAttribute(Globals.style + "row-height", $"{Height ?? "0.452cm"}"),
                    new XAttribute(Globals.fo + "break-before", "auto"),
                    new XAttribute(Globals.style + "use-optimal-row-height", $"{optimalRowHeight}")
                    ));

                Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "automatic-styles").Add(rowStyleElement);
            }
        }
    }
}