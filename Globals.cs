using System.Xml.Linq;

namespace ODSFiles
{
    internal static class Globals
    {
        public static XDocument content;
        public static XDocument styles;
        public static XDocument settings;

        public static readonly XNamespace office = @"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        public static readonly XNamespace table = @"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        public static readonly XNamespace text = @"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        public static readonly XNamespace style = @"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        public static readonly XNamespace calcext = @"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
        public static readonly XNamespace fo = @"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        public static readonly XNamespace number = @"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
        public static readonly XNamespace config = @"urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    }
}