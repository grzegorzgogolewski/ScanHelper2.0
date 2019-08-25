using System.Collections.Generic;

namespace ScanHelper
{
    public class KdokRodz
    {
        public int IdRodzDok { get; set; }
        public string Opis { get; set; }
        public string Prefix { get; set; }
        public bool Scal { get; set; }  //  czy scalać dokumenty danego rodzaju
        public int Count { get; set; }  //  liczba dokumentów danego typu
    }

    public class KdokRodzDict : Dictionary<int, KdokRodz>
    {

    }
}
