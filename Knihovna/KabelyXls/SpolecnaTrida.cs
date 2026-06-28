#nullable disable

using System;
using System.Linq;

namespace ExcelGenerateSeznam
{
    /// <summary>
    /// Společná bázová třída pro entity s automatickým generováním Apid.
    /// </summary>
    public abstract class SpolecnaTrida
    {
        /// <summary>
        /// Identifikátor entity.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Unikátní API identifikátor (8 náhodných znaků).
        /// </summary>
        public string Apid { get; set; }

        /// <summary>
        /// Inicializuje novou instanci třídy <see cref="SpolecnaTrida"/> a vygeneruje Apid.
        /// </summary>
        protected SpolecnaTrida()
        {
            Apid = GenerujApid();
        }

        /// <summary>
        /// Generuje náhodný 8znakový alfanumerický řetězec.
        /// </summary>
        /// <returns>Náhodný 8znakový řetězec.</returns>
        private static string GenerujApid()
        {
            const string znaky = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            Random random = new Random();
            return new string(Enumerable.Repeat(znaky, 8)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
