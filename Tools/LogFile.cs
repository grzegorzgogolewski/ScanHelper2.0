using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace Tools
{
    public static class LogFile
    {
        public static void SaveMessage(string message, [CallerMemberName] string callerName = "")
        {
            using (StreamWriter str = new StreamWriter(new FileStream("ScanHelper.log", FileMode.Append), Encoding.UTF8))
            {
                str.WriteLine($"{DateTime.Now}\t{callerName.PadRight(25, ' ')}:\t{message}");
            }
        }
    }
}
