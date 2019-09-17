using System.Collections.Generic;

namespace ScanHelper
{
    public class ScanFile
    {
        public int IdFile { get; set; }
        public string PathAndFileName { get; set; }
        public string Path { get; set; }
        public string FileName { get; set; }
        public string Prefix { get; set; }
        public int TypeCounter { get; set; }
        public bool Merged { get; set; } = false;
        public byte[] PdfFile { get; set; }

    }

    public class ScanFileDict : Dictionary<int, ScanFile>
    {

    }
}
