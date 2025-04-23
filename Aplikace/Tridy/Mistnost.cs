using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
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

        [Display(Name = "Povrchová Úprava Podlahy")]
        public string PovrchováÚpravaPodlahy { get; set; } = string.Empty;

        [Display(Name = "Povrchová Úprava Stropu")]
        public string PovrchováÚpravaStropu { get; set; } = string.Empty;

        [Display(Name = "Povrchová Úprava Stěny")]
        public string PovrchováÚpravaStěny { get; set; } = string.Empty;

        /// <summary>Stavební objekt</summary>
        public string Objekt { get; set; } = string.Empty;

        [JsonIgnore]
        public static Dictionary<int, string> Sloupce => new() {
                {1, "Číslo" },
                {2, "Název" },
                {3, "Podlaží" },
                {4, "Komentáře" },
                {5, "Obvod" },
                {6, "Plocha" },
                {7, "Objem" },
                //{8, "PovrchováÚpravaPodlahy" },
                //{9, "PovrchováÚpravaStropu" },
                //{10, "PovrchováÚpravaStěny" },
                {8, "Objekt" },
                //{12, "Sloupce" },
            };
    }
}
