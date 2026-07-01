namespace Knihovna.KabelyXls
{
    /// <summary>
    /// Reprezentuje položku kabelu v seznamu kabelů pro export do Excelu.
    /// Dědí ze společné třídy <see cref="SpolecnaTrida"/>.
    /// </summary>
    public class KabelPolozka : SpolecnaTrida
    {
        /// <summary>
        /// Číslo položky (No.).
        /// </summary>
        public string Polozka { get; set; } = string.Empty;

        /// <summary>
        /// Revize kabelu.
        /// </summary>
        public string Revize { get; set; } = string.Empty;

        /// <summary>
        /// Označení kabelu (Cable No.).
        /// </summary>
        public string CisloKabelu { get; set; } = string.Empty;

        /// <summary>
        /// Typ kabelu (Cable type).
        /// </summary>
        public string KabelTyp { get; set; } = string.Empty;

        /// <summary>
        /// Průřez kabelu (mm2).
        /// </summary>
        public string Prurez { get; set; } = string.Empty;

        /// <summary>
        /// Délka kabelu (m).
        /// </summary>
        public string Delka { get; set; } = string.Empty;

        /// <summary>
        /// Zdrojové zařízení (From).
        /// </summary>
        public string ZeZarizeni { get; set; } = string.Empty;

        /// <summary>
        /// Ukončení na straně zdroje (Gland).
        /// </summary>
        public string UkonceniZe { get; set; } = string.Empty;

        /// <summary>
        /// Cílové zařízení (To).
        /// </summary>
        public string DoZarizeni { get; set; } = string.Empty;

        /// <summary>
        /// Ukončení na straně cíle (Gland).
        /// </summary>
        public string UkonceniDo { get; set; } = string.Empty;

        /// <summary>
        /// Poznámka (Remark).
        /// </summary>
        public string Poznamka { get; set; } = string.Empty;
    }
}
