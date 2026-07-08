﻿namespace Knihovna.KabelyXls
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
        public string Popis { get; set; } = string.Empty;

        /// <summary>
        /// Typ a velikost zařízení.
        /// </summary>
        public string TypVelikost { get; set; } = string.Empty;

        ///// <summary>
        ///// Počet kusů (Ks).
        ///// </summary>

        ///// <summary>
        ///// PID označení.
        ///// </summary>

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
        public string InstalovanyPi { get; set; } = string.Empty;

        /// <summary>
        /// Výpočtový činný výkon Pp [kW].
        /// </summary>
        public string VypoctovyPi { get; set; } = string.Empty;

        ///// <summary>
        ///// Parametr IVCHS.
        ///// </summary>

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

