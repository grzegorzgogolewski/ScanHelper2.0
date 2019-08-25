using License;

namespace ScanHelper
{
    /// <summary>
    /// Klasa zawierająca zmienne globalne
    /// </summary>
    public static class Global
    {
        // obiekt licencji programu
        public static MyLicense License = new MyLicense();

        // obiekt z listą rodzajów dokumentów
        public static KdokRodzDict DokDict = new KdokRodzDict();

        // obiekt z listą skanów
        public static ScanFileDict ScanFiles = new ScanFileDict();

        //  aktualny zoom dokumentu
        public static int Zoom;

        //  ostatni przetwarzany katalog
        public static string LastDirectory;

        //  czy wstawiać znak wodny do pliku
        public static bool Watermark;

        public static bool SaveRotation;

        public static int IdSelectedFile;
    }
}
