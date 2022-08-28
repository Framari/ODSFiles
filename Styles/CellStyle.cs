using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace ODSFiles.Style
{
    public class CellStyle
    {
        internal Text text = new Text();
        internal Paragrapgh paragrapgh = new Paragrapgh();
        internal Number number = new Number();

        internal string StyleName { get; set; }
        internal int HashCodeStyle { get; set; }


        #region Default cell props
        internal string DataStyleName { get; set; }
        #endregion

        #region Custom cell props
        public string VerticalAlign {internal get; set; }
        public string HorizontalAlign { internal get; set; }
        public string WrapOption { internal get; set; }
        public string CellColor { internal get; set; }
        #endregion


        public Text Text
        {
            get { return text; }
        }

        public Number Number
        {
            get { return number; }
        }

        public Paragrapgh Paragrapgh
        {
            get { return paragrapgh; }
        }

        internal void CreatElement(string type)
        {
            DataStyleName = "N0";

            if (type == "date")
            {
                DataStyleName = $@"N3";

                bool exist = false;

                IEnumerable<XElement> elements = Globals.content.Element(Globals.office + "document-content")
                    .Element(Globals.office + "automatic-styles").Elements(Globals.number + "date-style");

                foreach (var elem in elements)
                {
                    if (elem.Attribute(Globals.style + "name").Value == $"{DataStyleName}")
                    {
                        exist = true;
                        break;
                    }
                }

                if (!exist)
                {
                    XElement dateStyle = new XElement(Globals.number + "date-style",
                        new XAttribute(Globals.style + "name", $"{DataStyleName}"),
                        new XAttribute(Globals.number + "automatic-order", "true")
                    );

                    dateStyle.Add(new XElement(Globals.number + "day",
                        new XAttribute(Globals.number + "style", "long")));
                    dateStyle.Add(new XElement(Globals.number + "text", "."));
                    dateStyle.Add(new XElement(Globals.number + "month",
                        new XAttribute(Globals.number + "style", "long")));
                    dateStyle.Add(new XElement(Globals.number + "text", "."));
                    dateStyle.Add(new XElement(Globals.number + "year"));

                    Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "automatic-styles").Add(dateStyle);
                }
            }

            if (type == "currency")
            {
                string group = number.Grouping == "true"  || number.Grouping == null ? "1" : "0";
                string currencySymbol = number.CurrencySymbol != null ? number.CurrencySymbol : "₽";
                string code = $"2{number.DecimalPlaces ?? "3"}{number.DecimalPlaces ?? "3"}{group}{currencySymbol}";

                int hashCodeDataStyle = code.GetHashCode();

                DataStyleName = $@"N{hashCodeDataStyle}";

                bool exist = false;

                IEnumerable<XElement> elements = Globals.content.Element(Globals.office + "document-content")
                    .Element(Globals.office + "automatic-styles").Elements(Globals.number + "number-style");

                foreach (var elem in elements)
                {
                    if (elem.Attribute(Globals.style + "name").Value == $"{DataStyleName}")
                    {
                        exist = true;
                        break;
                    }
                }

                if (!exist)
                {
                    XElement numberStyle = new XElement(Globals.number + "number-style",
                        new XAttribute(Globals.style + "name", $"{DataStyleName}")
                    );

                    numberStyle.Add(new XElement(Globals.number + "number",
                            new XAttribute(Globals.number + "decimal-places", number.DecimalPlaces ?? "3"),
                            new XAttribute(Globals.number + "min-decimal-places", number.DecimalPlaces ?? "3"),
                            new XAttribute(Globals.number + "min-integer-digits", "1"),
                            new XAttribute(Globals.number + "grouping", number.Grouping ?? "true")
                        )
                    );

                    numberStyle.Add(new XElement(Globals.number + "text", $" {currencySymbol}"));

                    Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "automatic-styles").Add(numberStyle);
                }
            }

            if (type == "float")
            {
                string group = number.Grouping == "true" || number.Grouping == null ? "1" : "0";
                string code = $"3{number.DecimalPlaces ?? "3"}{number.DecimalPlaces ?? "3"}{group}";

                int hashCodeDataStyle = code.GetHashCode();

                DataStyleName = $@"N{hashCodeDataStyle}";

                bool exist = false;

                IEnumerable<XElement> elements = Globals.content.Element(Globals.office + "document-content")
                    .Element(Globals.office + "automatic-styles").Elements(Globals.number + "number-style");

                foreach (var elem in elements)
                {
                    if (elem.Attribute(Globals.style + "name").Value == $"{DataStyleName}")
                    {
                        exist = true;
                        break;
                    }
                }

                if (!exist)
                {
                    XElement numberStyle = new XElement(Globals.number + "number-style",
                        new XAttribute(Globals.style + "name", $"{DataStyleName}")
                    );

                    numberStyle.Add(new XElement(Globals.number + "number",
                            new XAttribute(Globals.number + "decimal-places", number.DecimalPlaces ?? "3"),
                            new XAttribute(Globals.number + "min-decimal-places", number.DecimalPlaces ?? "3"),
                            new XAttribute(Globals.number + "min-integer-digits", "1"),
                            new XAttribute(Globals.number + "grouping", number.Grouping ?? "true")
                        )
                    );

                    numberStyle.Add(new XElement(Globals.number + "text", " "));

                    Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "automatic-styles").Add(numberStyle);
                }
            }

            #region CellParameters

            XAttribute verticalAlign = VerticalAlign != null
                ? new XAttribute(Globals.style + "vertical-align", $"{VerticalAlign}")
                : null;

            XAttribute wrapOption = WrapOption != null
                ? new XAttribute(Globals.fo + "wrap-option", $"{WrapOption}")
                : null;

            XAttribute cellColor = CellColor != null
                ? new XAttribute(Globals.fo + "background-color", $"{CellColor}")
                : null;

            XElement cellProperties = verticalAlign != null || wrapOption != null || cellColor != null
                ? new XElement(Globals.style + "table-cell-properties", verticalAlign, wrapOption, cellColor)
                : null;

            #endregion

            #region CellParagraph

            XAttribute horizontalAlign = HorizontalAlign != null
                ? new XAttribute(Globals.fo + "text-align", $"{HorizontalAlign}")
                : null;

            XAttribute marginLeft = paragrapgh.MarginLeft != null
                ? new XAttribute(Globals.fo + "margin-left", $"{paragrapgh.MarginLeft}")
                : null;

            XAttribute marginRight = paragrapgh.MarginLeft != null
                ? new XAttribute(Globals.fo + "margin-right", $"{paragrapgh.MarginRight}")
                : null;

            XElement paragraphProperties = horizontalAlign != null || marginLeft !=null || marginRight != null
                ? new XElement(Globals.style + "paragraph-properties", horizontalAlign, marginLeft, marginRight) 
                : null;

            #endregion

            #region CellText

            XAttribute textSize = text.TextSize != null
                ? new XAttribute(Globals.fo + "font-size", text.TextSize)
                : null;
            XAttribute textSizeAsian = textSize != null
                ? new XAttribute(Globals.style + "font-size-asian", text.TextSize)
                : null;
            XAttribute textSizeComplex = textSize != null
                ? new XAttribute(Globals.style + "font-size-complex", text.TextSize)
                : null;

            XAttribute textColor = text.TextColor != null
                ? new XAttribute(Globals.fo + "color", $"{text.TextColor}")
                : null;

            XAttribute fontStyle = text.Italic == "true"
                ? new XAttribute(Globals.fo + "font-style", "italic")
                : null;
            XAttribute fontStyleAsian = fontStyle != null
                ? new XAttribute(Globals.style + "font-style-asian", "italic")
                : null;
            XAttribute fontStyleComplex = fontStyle != null
                ? new XAttribute(Globals.style + "font-style-complex", "italic")
                : null;

            XAttribute textUnderlineStyle = text.Solid == "true"
                ? new XAttribute(Globals.style + "text-underline-style", "solid")
                : null;
            XAttribute textUnderlineWidth = textUnderlineStyle != null
                ? new XAttribute(Globals.style + "text-underline-width", "auto")
                : null;
            XAttribute textUnderlineColor = textUnderlineStyle!= null
                ? new XAttribute(Globals.style + "text-underline-color", "font-color")
                : null;

            XAttribute fontWeight = text.Bold == "true"
                ? new XAttribute(Globals.fo + "font-weight", "bold")
                : null;
            XAttribute fontWeightAsian = fontWeight != null
                ? new XAttribute(Globals.style + "font-weight-asian", "bold")
                : null;
            XAttribute fontWeightComplex = fontWeight != null
                ? new XAttribute(Globals.style + "font-weight-complex", "bold")
                : null;

            XElement textProperties = textColor!=null || textSize != null || fontStyle!=null || fontWeight != null || textUnderlineStyle !=null || fontWeight != null
                ? new XElement(Globals.style + "text-properties",
                    textColor,
                    textSize, textSizeAsian, textSizeComplex,
                    fontStyle, fontStyleAsian, fontStyleComplex,
                    textUnderlineStyle, textUnderlineWidth, textUnderlineColor,
                    fontWeight, fontWeightAsian, fontWeightComplex
                    )
                : null;

            #endregion

            StyleName = $"{verticalAlign}{wrapOption}{cellColor}{horizontalAlign}{marginLeft}{marginRight}{textColor}{textSize}{fontStyle}{textUnderlineStyle}{fontWeight}{DataStyleName}";
            HashCodeStyle = StyleName.GetHashCode();

            bool styleExist = false;

            IEnumerable<XElement> styleElements = Globals.content.Element(Globals.office + "document-content")
                .Element(Globals.office + "automatic-styles").Elements(Globals.style + "style");

            foreach (var elem in styleElements)
            {
                if (elem.Attribute(Globals.style + "name").Value == $"ce{HashCodeStyle}")
                {
                    styleExist = true;
                    break;
                }
            }

            if (!styleExist)
            {
                XElement cellStyleElement = new XElement(Globals.style + "style",
                    new XAttribute(Globals.style + "name", $"ce{HashCodeStyle}"),
                    new XAttribute(Globals.style + "family", "table-cell"),
                    new XAttribute(Globals.style + "parent-style-name", "Default"),
                    new XAttribute(Globals.style + "data-style-name", $"{DataStyleName}"),
                    cellProperties,
                    paragraphProperties,
                    textProperties
                    );

                Globals.content.Element(Globals.office + "document-content").Element(Globals.office + "automatic-styles").Add(cellStyleElement);
            }
        }
    }
}