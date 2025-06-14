using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy {
    public class Jistic {
        public string Druh { get; set; } = string.Empty;
        public string Velikost { get; set; } = string.Empty;
        public double Icu { get; set; }
        public double In { get; set; }
        public string ObjednaciKod { get; set; } = string.Empty;
        public string Ir { get; set; }= string.Empty;
        public double Irmin { get; set; }
        public double Irmax { get; set; }
        public string Ii { get; set; }= string.Empty;
        public double Iimin { get; set; }
        public double Iimax { get; set; }
        public string Isd { get; set; }= string.Empty;
        public double Isdmin { get; set; }
        public double Isdmax { get; set; }
        public string Spoust { get; set; } = string.Empty;
        public double Hmotnost { get; set; }

        // Rozměry
        public string Rozmery { get; set; } = string.Empty;
        public double Sirka { get; set; }
        public double Vyska { get; set; }
        public double Hloubka { get; set; }
    }
}
