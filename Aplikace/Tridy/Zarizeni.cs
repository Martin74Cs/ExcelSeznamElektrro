using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Zarizeni
    {
        public string Tag { get; set; } = string.Empty;

        /// <summary>Jméno zařízení</summary>
        [Display(Name = "Jméno zařízení")]
        public string Popis { get; set; } = string.Empty;
        public string Prikon { get; set; } = string.Empty;

        public string PID { get; set; } = string.Empty;

        public double HP => double.TryParse(Prikon, out double hodnota) ? hodnota * 3.29 : 0; // Převod textu na číslo a součet

        public string Druh { get; set; } = string.Empty;

        public string Menic { get; set; } = string.Empty;

        public string BalenaJednotka { get; set; } = string.Empty;

        public string Proud { get; set; } = string.Empty;
        public string Delka { get; set; } = string.Empty;

        public string AWG { get; set; } = string.Empty;

        /// <summary>Rozvaděč</summary>
        public string Rozvadec { get; set; } = string.Empty;

        public string RozvadecCislo { get; set; } = string.Empty;

        public string PruzezMM2 { get; set; } = string.Empty;

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

        //public static void NastavVlastnost(string nazevVlastnosti, object hodnota)
        public static void NastavVlastnost(object obj, string nazevVlastnosti, object hodnota)
        {
            // Získáme typ objektu
            //var typ = obj.GetType();
            //object obj = new Zarizeni();
            var typ = obj.GetType();
            // Získáme vlastnost podle názvu
            var vlastnost = typ.GetProperty(nazevVlastnosti);

            if (vlastnost != null && vlastnost.CanWrite)
            {
                // Nastavíme hodnotu vlastnosti
                vlastnost.SetValue(obj, Convert.ChangeType(hodnota, vlastnost.PropertyType));
            }
            else
            {
                throw new ArgumentException($"Vlastnost {nazevVlastnosti} neexistuje nebo není zapisovatelná.");
            }
        }

    }
}
