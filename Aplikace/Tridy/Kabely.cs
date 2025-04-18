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
        //public static string[] PoleVstup = ["Jmeno", "Vek", "Mesto"];

        public static void Vypis(List<Zarizeni> zaznamy)
        {
            // Vypsání hodnot záznamů podle názvů parametrů
            foreach (var zaznam in zaznamy)
            {
                foreach (var Property in zaznam.GetType().GetProperties())
                {
                    // Pomocí reflexe získáme hodnotu vlastnosti
                    //PropertyInfo vlastnost = zaznam.GetType().GetProperty(Property);
                    if (Property != null)
                    {
                        var hodnota = Property.GetValue(zaznam);
                        Console.WriteLine($"{Property.Name}: {hodnota}");
                    }
                }
                Console.WriteLine();
            }
        }

    }
}
