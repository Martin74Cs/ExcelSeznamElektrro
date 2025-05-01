using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Motor
    {
        [Display(Name = "Výrobce motoru")]
        [Required(ErrorMessage = "Hodnota kategorie je vyžadována")]
        public string Vyrobce { get; set; } = string.Empty;

        [Display(Name = "Typ motoru")]
        public string TypMotoru { get; set; } = string.Empty;

        [Display(Name = "Druh materiálu")]
        public string Material { get; set; } = string.Empty;

        [Display(Name = "Jmenovitý výkon při 50 Hz Pn50 [kW]")]
        public double Vykon50 { get; set; }

        [Display(Name = "Jmenovitý výkon při 60 Hz Pn60 [kW]")]
        public double Vykon60 { get; set; }

        [Display(Name = "Velikost")]
        public string Velikost { get; set; } = string.Empty;

        [Display(Name = "Jmenovitý otáčky při 50Hz nN [1/min]")]
        public int Otacky50 { get; set; }

        [Display(Name = "Jmenovitý moment při 50Hz MN [Nm]")]
        public double Moment50 { get; set; }

        [Display(Name = "Účinosti 4/4 nN [%]")]
        public double Ucinnost44 { get; set; }

        [Display(Name = "Účinosti 3/4 nN [%]")]
        public double Ucinnost34 { get; set; }

        [Display(Name = "Účinosti 2/4 nN [%]")]
        public double Ucinnost24 { get; set; }

        [Display(Name = "Účiník při 50 Hz 4/4 nN [%] Cos(fi)")]
        public double Ucinik50 { get; set; }

        [Display(Name = "Jmenovitý proud při 400V [A]")]
        public double Proud400 { get; set; }

        [Display(Name = "Záběrný moment Ma/Mn [-]")]
        public double MaMn { get; set; }

        [Display(Name = "Záběrný proud Ia/In [-]")]
        public double IaIn { get; set; }

        [Display(Name = "Maximální moment Mk/Mn [-]")]
        public double MkMn { get; set; }

        [Display(Name = "Hladina akustického tlaku při 50 Hz Lpfa [dB(A)]")]
        public double Lpfa { get; set; }

        [Display(Name = "Hladina akustického výkonu při 50 Hz Lwa [db(A)]")]
        public double LWA { get; set; }

        [Display(Name = "Objednací číslo Trojůhelnik/Hvězda [V]")]
        public string ObjednaciCislo { get; set; } = string.Empty;

        [Display(Name = "Popis Trojůhelnik/Hvězda [V]")]
        public string TrojuhelnikHvezda { get; set; } = string.Empty;

        [Display(Name = "Tvar připojení")]
        public string Priruba { get; set; } = string.Empty;

        [Display(Name = "Ochrana termistorů")]
        public string Termistory { get; set; } = string.Empty;

        [Display(Name = "Hmotnost při tvaru Mimb3 [kg]")]
        public double Hmotnost { get; set; }

        [Display(Name = "Moment setrvačnosti J [kgm3]")]
        public double MomentSetrvacnosti { get; set; }

        [Display(Name = "Momentová třída")]
        public double MomentovaTrida { get; set; }

        [Display(Name = "Počet polů")]
        public int Poly { get; set; }
    }

    //Doplneni dalsich hodnot
    public class MotoryDalsi : Motor
    {
        [Display(Name = "Nazev motoru")]
        public string Nazev { get; set; } = string.Empty;

        [Display(Name = "Jmenovitý otáčky při 60Hz nN [1/min]")]
        public int Otacky60 { get; set; }

        [Display(Name = "Otáčky snížené o skluz při 50Hz nN [1/min]")]
        public int Otacky { get; set; }

        ///<summary>způsob chlazení IC 411 podle  ČSN EN 60034-6</summary>
        [Display(Name = "Chlazení")]
        public string Chlazeni { get; set; } = string.Empty;

        [Display(Name = "Popis Chlazení")]
        public string PopisChlazeni { get; set; } = string.Empty;

        [Display(Name = "Tepelní třída")]
        public string TeplotniTrida { get; set; } = string.Empty;

        [Display(Name = "Třída účinosti IE [-]")]
        public string Trida50 { get; set; }= string.Empty;

        [Display(Name = "Třída účinosti IE [-]")]
        public string Trida60 { get; set; }= string.Empty;

        [Display(Name = "Material kostry")]
        public string Kostra { get; set; } = string.Empty;

        [Display(Name = "Krytí")]
        public int Kryti { get; set; }
    }
}

