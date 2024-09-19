using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Kabely
    {
        public string Tag { get; set; } = string.Empty;
        /// <summary>Jméno zařízení</summary>
        public string MCC { get; set; } = string.Empty;
        public string CisloMCC { get; set; } = string.Empty;
        public string Oznaceni { get; set; } = string.Empty;

        public string Kabel { get; set; } = string.Empty;
        public string PocetZil { get; set; } = string.Empty;
        public string Prurezmm2 { get; set; } = string.Empty;
        public string PrurezFt { get; set; } = string.Empty;

        public string Druh { get; set; } = string.Empty;

        public string OdkudSvokra { get; set; } = string.Empty;
        
        public string Delka { get; set; } = string.Empty;
        /// <summary>Rozvaděč</summary>

        // Seznam názvů parametrů (vlastností), které chceme vypsat
        public static List<string> PoleVstup = ["Jmeno", "Vek", "Mesto"];

        public static void Vypis(List<Zarizeni> zaznamy)
        {
            // Vypsání hodnot záznamů podle názvů parametrů
            foreach (var zaznam in zaznamy)
            {
                foreach (var nazevParametru in PoleVstup)
                {
                    // Pomocí reflexe získáme hodnotu vlastnosti
                    PropertyInfo vlastnost = zaznam.GetType().GetProperty(nazevParametru);
                    if (vlastnost != null)
                    {
                        var hodnota = vlastnost.GetValue(zaznam);
                        Console.WriteLine($"{nazevParametru}: {hodnota}");
                    }
                }
                Console.WriteLine();
            }
        }

    }
}
