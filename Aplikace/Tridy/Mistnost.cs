using Aplikace.Rozšíření;
using Aplikace.Upravy;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Mistnost : Entity
    {

        [Display(Name = "Číslo místnosti")]
        public string Číslo { get; set; } = string.Empty;

        [Display(Name = "Název místnosti")]
        public string Název { get; set; } = string.Empty;

        [Display(Name = "Obvod místnosti")]
        public string Podlaží { get; set; } = string.Empty;
        public string Komentáře { get; set; } = string.Empty;

        [Display(Name = "Obvod místnosti")]
        public double Obvod { get; set; }

        [Display(Name = "Plocha mísnosti")]
        public double Plocha { get; set; }

        [Display(Name = "Objem místnosti")]
        public double Objem { get; set; }

        [Display(Name = "Povrchová Úprava Podlahy")]
        public string PovrchováÚpravaPodlahy { get; set; } = string.Empty;

        [Display(Name = "Povrchová Úprava Stropu")]
        public string PovrchováÚpravaStropu { get; set; } = string.Empty;

        [Display(Name = "Povrchová Úprava Stěny")]
        public string PovrchováÚpravaStěny { get; set; } = string.Empty;

        /// <summary>Stavební objekt</summary>
        [Display(Name = "Stavební objekt")]
        public string Objekt { get; set; } = string.Empty;

        [JsonIgnore]
        internal static string[] Nadpis =>
        [
            "Apid",
            "Objekt",
            "Číslo",
            "Název",
            "Podlaží",
            "Komentáře",
            "Obvod",
            "Plocha",
            "Objem",
            // "PovrchováÚpravaPodlahy",
            // "PovrchováÚpravaStropu",
            // "PovrchováÚpravaStěny",
        ];

        [JsonIgnore]
        //public static Dictionary<int, string> Sloupce => new() {
        //        {1, "Objekt" },
        //        {2, "Číslo" },
        //        {2, "Název" },
        //        {3, "Podlaží" },
        //        {4, "Komentáře" },
        //        {5, "Obvod" },
        //        {6, "Plocha" },
        //        {7, "Objem" },
        //        //{8, "PovrchováÚpravaPodlahy" },
        //        //{9, "PovrchováÚpravaStropu" },
        //        //{10, "PovrchováÚpravaStěny" },
        //        //{12, "Sloupce" },
        //    };
        public static IDictionary<int, string> Sloupce => Nadpis
            .Select((name, index) => new { Index = index + 1, Name = name })
            .ToDictionary(x => x.Index, x => x.Name);
    }

    public class Slaboproudy : Mistnost
    {
        public static Slaboproudy CopyToParent(Mistnost child)
        {
            var parent = new Slaboproudy();
            var childProps = typeof(Mistnost).GetProperties();
    
            foreach (var parentProp in typeof(Slaboproudy).GetProperties())
            {
         
                // Najdeme odpovídající vlastnost v Child
                var childProp = childProps.FirstOrDefault(p => p.Name == parentProp.Name && p.PropertyType == parentProp.PropertyType);
                if (childProp != null)
                {
                    if (childProp.Name == "Sloupce") continue;
                    var value = childProp.GetValue(child);
                    parentProp.SetValue(parent, value);
                }
            }

            return parent;
        }

        //[Display(Name = "Pokus")]
        //public string Pokus { get; set; } = string.Empty;
    
        [Display(Name = "Hlásič")]
        public string EpsHlasic { get; set; } = string.Empty;

        [Display(Name = "Siréna")]
        public string EpsSirena { get; set; } = string.Empty;
        
        [Display(Name = "Plynový detekce")]
        public string GDS { get; set; } = string.Empty;

        [Display(Name = "Ethernet")]
        public string Ethernet { get; set; } = string.Empty;

        [Display(Name = "Kamerový systém")]
        public string Kamera { get; set; } = string.Empty;

        [Display(Name = "Přístupový systém")]
        public string ACS { get; set; } = string.Empty;

        [JsonIgnore]
        internal static string[] Nadpis =>
        [
            "EpsHlasic",
            "EpsSirena",
            "GDS",
            "Ethernet",
            "Kamera",
            "ACS",
            "Pokus"
        ];

        [JsonIgnore]
        public static IDictionary<int, string> Sloupce => Nadpis
            .Select((name, index) => new { Index = index + 1, Name = name })
            .ToDictionary(x => x.Index, x => x.Name);


        [JsonIgnore]
        /// <summary>Sloupce pro zobrazení v tabulce, ze seznamu vytvoženy IDictionary čísla označují sloupce</summary>
        public static IDictionary<int, string> SloupceSpojit => Mistnost.Nadpis
                .Concat(Nadpis)
                .Select((name, index) => new { Index = index + 1, Name = name })
                .ToDictionary(x => x.Index, x => x.Name);

        //=> Sloupce.AddRange(Mistnost.Sloupce);
    }
}
