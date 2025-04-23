using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Mistnost
    {
        public string Číslo { get; set; } = string.Empty;
        public string Název { get; set; } = string.Empty;
        public string Podlaží { get; set; } = string.Empty;
        public string Komentáře { get; set; } = string.Empty;

        public double Obvod { get; set; }
        public double Plocha { get; set; }
        public double Objem { get; set; }

        public string PovrchováÚpravaPodlahy { get; set; } = string.Empty;
        public string PovrchováÚpravaStropu { get; set; } = string.Empty;
        public string PovrchováÚpravaStěny { get; set; } = string.Empty;

        /// <summary>Stavební objekt</summary>
        public string Objekt { get; set; } = string.Empty;
    }
}
