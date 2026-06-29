#nullable disable

using System.Collections.Generic;

namespace ExcelGenerateSeznam
{
    /// <summary>
    /// Reprezentuje data pro vyplnění úvodního listu Cover.
    /// </summary>
    public class CoverData
    {
        /// <summary>
        /// Název zákazníka (ZAKAZNIK -> Cover!$C$18).
        /// </summary>
        public string Zakaznik { get; set; } = string.Empty;

        /// <summary>
        /// Číslo nebo kód projektu (PROJEKT -> Cover!$C$13).
        /// </summary>
        public string Projekt { get; set; } = string.Empty;

        /// <summary>
        /// Název akce / projektu (NAZEV -> Cover!$C$19).
        /// </summary>
        public string Nazev { get; set; } = string.Empty;

        /// <summary>
        /// Název dokumentu (např. SOUPIS SPOTŘEBIČŮ ELEKTRO -> _NA4 -> Cover!$C$20).
        /// </summary>
        public string DokumentNazev { get; set; } = string.Empty;

        /// <summary>
        /// Název části technologie (např. TECHNOLOGICKÁ ELEKTROINSTALACE -> _CAS4 -> Cover!$C$12).
        /// </summary>
        public string Technologie { get; set; } = string.Empty;

        /// <summary>
        /// Typ dokumentu (např. TP-N- -> CDOK1 -> Cover!$D$9).
        /// </summary>
        public string CistyDokumentTyp { get; set; } = string.Empty;

        /// <summary>
        /// Číslo dokumentu (např. 9446 -> CDOK2 -> Cover!$E$9).
        /// </summary>
        public string CistyDokumentCislo { get; set; } = string.Empty;

        /// <summary>
        /// Aktuální revize (např. A -> REV -> Cover!$F$9).
        /// </summary>
        public string Revize { get; set; } = string.Empty;

        /// <summary>
        /// Seznam revizí pro revizní tabulku (max 6 řádků).
        /// </summary>
        public List<RevizeInfo> RevizeSeznam { get; set; } = [];
    }

    /// <summary>
    /// Reprezentuje jeden řádek v revizní tabulce na úvodním listu.
    /// </summary>
    public class RevizeInfo
    {
        /// <summary>
        /// Označení revize (např. A, B, 0).
        /// </summary>
        public string Rev { get; set; } = string.Empty;

        /// <summary>
        /// Datum revize (např. 24.06.2026).
        /// </summary>
        public string Date { get; set; } = string.Empty;

        /// <summary>
        /// Popis změn v revizi.
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// Status revize (např. DFT, PRL, FIN, atd.).
        /// </summary>
        public string Stat { get; set; } = string.Empty;

        /// <summary>
        /// Jméno osoby, která revizi vypracovala.
        /// </summary>
        public string Zpacoval { get; set; } = string.Empty;

        /// <summary>
        /// Jméno osoby, která revizi zkontrolovala.
        /// </summary>
        public string Kontroloval { get; set; } = string.Empty;

        /// <summary>
        /// Jméno osoby, která revizi schválila.
        /// </summary>
        public string Schvalil { get; set; } = string.Empty;
    }
}
