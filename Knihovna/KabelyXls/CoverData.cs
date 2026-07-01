using System.Collections.Generic;

namespace Knihovna.KabelyXls
{
    /// <summary>
    /// Reprezentuje data pro vyplnění úvodního listu Cover.
    /// Dědí ze společné třídy <see cref="SpolecnaTrida"/>.
    /// </summary>
    public class CoverData : SpolecnaTrida
    {
        /// <summary>
        /// Název zákazníka (ZAKAZNIK).
        /// </summary>
        public string Zakaznik { get; set; } = string.Empty;

        /// <summary>
        /// Číslo nebo kód projektu (PROJEKT).
        /// </summary>
        public string Projekt { get; set; } = string.Empty;

        /// <summary>
        /// Název akce / projektu (NAZEV).
        /// </summary>
        public string Nazev { get; set; } = string.Empty;

        /// <summary>
        /// Název dokumentu.
        /// </summary>
        public string DokumentNazev { get; set; } = string.Empty;

        /// <summary>
        /// Název části technologie.
        /// </summary>
        public string Technologie { get; set; } = string.Empty;

        /// <summary>
        /// Typ dokumentu.
        /// </summary>
        public string CistyDokumentTyp { get; set; } = string.Empty;

        /// <summary>
        /// Číslo dokumentu.
        /// </summary>
        public string CistyDokumentCislo { get; set; } = string.Empty;

        /// <summary>
        /// Aktuální revize.
        /// </summary>
        public string Revize { get; set; } = string.Empty;

        /// <summary>
        /// Seznam revizí pro revizní tabulku (max 6 řádků).
        /// </summary>
        public List<RevizeInfo> RevizeSeznam { get; set; } = [];
    }
}
