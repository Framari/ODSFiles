using System;
using System.Xml.Linq;
using System.IO;
using Ionic.Zip;

namespace ODSFiles
{
    public class Document
    {
        private ZipFile _zip = new ZipFile();
        private MemoryStream _content = new MemoryStream();
        private MemoryStream _styles = new MemoryStream();
        private MemoryStream _settings = new MemoryStream();

        private Sheets _sheets = new Sheets();
        

        public Sheets Sheet
        {
            get { return _sheets; }
        }

        public Document()
        {
            var a = System.Reflection.Assembly.GetExecutingAssembly();

            _zip.AddEntry("META-INF/manifest.xml", a.GetManifestResourceStream("ODSFiles.Resorces.pattern.META_INF.manifest.xml"));
            _zip.AddEntry("content.xml", a.GetManifestResourceStream("ODSFiles.Resorces.pattern.content.xml"));
            _zip.AddEntry("meta.xml", a.GetManifestResourceStream("ODSFiles.Resorces.pattern.meta.xml"));
            _zip.AddEntry("mimetype", a.GetManifestResourceStream("ODSFiles.Resorces.pattern.mimetype"));
            _zip.AddEntry("settings.xml", a.GetManifestResourceStream("ODSFiles.Resorces.pattern.settings.xml"));
            _zip.AddEntry("styles.xml", a.GetManifestResourceStream("ODSFiles.Resorces.pattern.styles.xml"));

            Globals.content = XDocument.Load(a.GetManifestResourceStream("ODSFiles.Resorces.pattern.content.xml"));
            Globals.styles = XDocument.Load(a.GetManifestResourceStream("ODSFiles.Resorces.pattern.styles.xml"));
            Globals.settings = XDocument.Load(a.GetManifestResourceStream("ODSFiles.Resorces.pattern.settings.xml"));
        }

        public void Save(FileInfo fileName)
        {
            Globals.content.Save(_content);
            Globals.styles.Save(_styles);
            Globals.settings.Save(_settings);

            _zip.UpdateEntry("content.xml", _content.ToArray());
            _zip.UpdateEntry("styles.xml", _styles.ToArray());
            _zip.UpdateEntry("settings.xml", _settings.ToArray());

            _zip.Save(fileName.DirectoryName + $@"\{fileName.Name + ".ods"}");
        }
    }
}
