#nullable disable

namespace ExcelGenerateSeznam
{
    /// <summary>
    /// Reprezentuje spotřebič elektro v soupisu spotřebičů.
    /// Dědí ze společné třídy <see cref="SpolecnaTrida"/>.
    /// </summary>
    public class Spotrebic : SpolecnaTrida
    {
        /// <summary>
        /// Číslo položky (No.).
        /// </summary>
        public string Polozka { get; set; } = string.Empty;

        /// <summary>
        /// Revize položky.
        /// </summary>
        public string Rev { get; set; } = string.Empty;

        /// <summary>
        /// Prostorové uspořádání (PU).
        /// </summary>
        public string BalenaJednotka { get; set; } = string.Empty;

        /// <summary>
        /// Umístění spotřebiče (např. místnost, patro).
        /// </summary>
        public string Umisteni { get; set; } = string.Empty;

        /// <summary>
        /// Technologické označení (tag).
        /// </summary>
        public string Tag { get; set; } = string.Empty;

        /// <summary>
        /// Název nebo popis zařízení.
        /// </summary>
        public string Zarizeni { get; set; } = string.Empty;

        /// <summary>
        /// Typ a velikost zařízení.
        /// </summary>
        public string TypVelikost { get; set; } = string.Empty;

        /// <summary>
        /// Počet kusů (Ks).
        /// </summary>
        public string Ks { get; set; } = string.Empty;

        /// <summary>
        /// PID označení.
        /// </summary>
        public string Pid { get; set; } = string.Empty;

        /// <summary>
        /// Napájecí rozváděč.
        /// </summary>
        public string Rozvadec { get; set; } = string.Empty;

        /// <summary>
        /// Jmenovité napětí (např. 400, 230).
        /// </summary>
        public string Napeti { get; set; } = string.Empty;

        /// <summary>
        /// Instalovaný činný výkon Pi [kW].
        /// </summary>
        public double? InstalovanyPi { get; set; }

        /// <summary>
        /// Výpočtový činný výkon Pp [kW].
        /// </summary>
        public double? VypoctovyPi { get; set; }

        /// <summary>
        /// Parametr IVCHS.
        /// </summary>
        public string Ivchs { get; set; } = string.Empty;

        /// <summary>
        /// Způsob startu motoru.
        /// </summary>
        public string StartMotoru { get; set; } = string.Empty;

        /// <summary>
        /// Doplňující poznámka.
        /// </summary>
        public string Poznamka { get; set; } = string.Empty;
    }
}
