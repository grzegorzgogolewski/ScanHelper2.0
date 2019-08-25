using System.Collections.Generic;

namespace ScanHelper
{
    public class ScanFile
    {
        public int IdFile { get; set; }
        public string PathAndFileName { get; set; }
        public string Path { get; set; }
        public string FileName { get; set; }
        public System.Drawing.Image ImageFile { get; set; }
        public byte[] PdfFile { get; set; }
        public string Prefix { get; set; }
    }

    public class ScanFileDict : Dictionary<int, ScanFile>
    {

    }
}
