using System.IO;
using IniParser;
using IniParser.Model;

namespace Tools
{
    public static class IniSettings
    {
        public static void SaveIni(string section, string key, string value)
        {
            if (!File.Exists("ScanHelper.ini")) File.Create("ScanHelper.ini").Dispose();

            FileIniDataParser parser = new FileIniDataParser();
            IniData iniFile = parser.ReadFile("ScanHelper.ini");
            iniFile[section][key] = value;

            parser.WriteFile("ScanHelper.ini", iniFile);
        }

        public static string ReadIni(string section, string key)
        {
            if (!File.Exists("ScanHelper.ini")) File.Create("ScanHelper.ini").Dispose();

            FileIniDataParser parser = new FileIniDataParser();
            IniData iniFile = parser.ReadFile("ScanHelper.ini");
            return iniFile[section][key];
        }

        public static void SaveDefaults()
        {
            SaveIni("Options", "Watermark", "0");
            SaveIni("Options", "SaveRotation", "0");

            SaveIni("FormMain", "x", "0");
            SaveIni("FormMain", "y", "0");
        }
    }
}