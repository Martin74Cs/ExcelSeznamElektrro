using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public  class Stykac
    {
        // Příkon ve W nebo kW (není úplně jasné, jaká jednotka – přizpůsob dle potřeby)
        public double Prikon { get; set; }

        // Proud v režimu AC1
        public double ProudAC1 { get; set; }

        // Výkon ve HP při 480 V
        public double Prikon480Hp { get; set; }

        // Proud v AC režimu (obecně)
        public int ProudAC { get; set; }

        // Jmenovité napětí
        public int Napeti { get; set; }

        // Rozsah napětí ovládání – AC
        public string NapetiOvladaniAC { get; set; } = string.Empty; 

        // Rozsah napětí ovládání – DC
        public string NapetiOvladaniDC { get; set; } = string.Empty; 

        // Počet normálně otevřených kontaktů (NO)
        public int NO { get; set; }

        // Počet normálně zavřených kontaktů (NC)
        public int NC { get; set; }

        // Typ stykače
        public string Typ { get; set; } = string.Empty;

        // Objednací kód
        public string ObjednaciKod { get; set; } = string.Empty;

        // Hmotnost v kg (nebo tunách – ale dle hodnoty to bude spíš kg)
        public double Hmotnost { get; set; }

        // Maximální teplota (např. okolí) v °C
        public int Teplota { get; set; }
        public string Výrobce { get; set; }
    }
}
