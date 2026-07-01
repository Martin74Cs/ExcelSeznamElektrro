using System;

namespace Knihovna.KabelyXls
{
    /// <summary>
    /// Reprezentuje jeden řádek v revizní tabulce na úvodním listu.
    /// Dědí ze společné třídy <see cref="SpolecnaTrida"/>.
    /// </summary>
    public class RevizeInfo : SpolecnaTrida
    {
        /// <summary>
        /// Označení revize (např. A, B, 0).
        /// </summary>
        public string Revize { get; set; } = string.Empty;

        /// <summary>
        /// Datum revize (např. 24.06.2026).
        /// </summary>
        public string Datum { get; set; } = string.Empty;

        /// <summary>
        /// Popis změn v revizi.
        /// </summary>
        public string PopisRevize { get; set; } = string.Empty;

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
