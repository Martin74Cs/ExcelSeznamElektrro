using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{

    public class Zarizeni : Entity
    {
        public Zarizeni() { }
        public Zarizeni(string Tag, string PID, string Popis, string Prikon ,string BalenaJednotka, string Menic , string Nic, string HPstr , string Proud, string PruzezMM2, string AWG, 
        string Delka, string Delkaft , string Rozvadec , string RozvadecCislo )
        {
            this.Tag = Tag;
            this.PID = PID;
            this.Popis = Popis;
            this.Prikon = Prikon;
            this.BalenaJednotka = BalenaJednotka;
            this.Menic = Menic;
            this.Nic = Nic;
            //hp
            this.Proud = Proud;
            this.PruzezMM2 = PruzezMM2;
            this.AWG = AWG;
            this.Delka = Delka;
            this.Delkaft = Delkaft;
            this.Rozvadec = Rozvadec;
            this.RozvadecCislo = RozvadecCislo;
            //this.Druh = Druh;
        }
        /// <summary>Druh zařízení čerpadlo, motor, trafo</summary>

        //var TextPole = new string[] { "Tag", "PID", "Equipment name", "kW", "BalenaJednotka", "Menic", "Nic", "Power [HP]", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
        public string Tag { get; set; } = string.Empty;
        public string PID { get; set; } = string.Empty;

        /// <summary>Jméno zařízení</summary>
        [Display(Name = "Jméno zařízení")]
        public string Popis { get; set; } = string.Empty;
        public string Prikon { get; set; } = string.Empty;
        public string BalenaJednotka { get; set; } = string.Empty;
        public string Menic { get; set; } = string.Empty;
        public string Nic { get; set; } = string.Empty;
        public double HP => double.TryParse(Prikon, out double hodnota) ? hodnota * 1.341022 : 0; // Převod textu na číslo a na koně
        public string Proud { get; set; } = string.Empty;
        public string PruzezMM2 { get; set; } = string.Empty;
        public string AWG { get; set; } = string.Empty;
        public string Delka { get; set; } = "100";
        public string Delkaft { get; set; } = string.Empty;

        public string Rozvadec { get; set; } = string.Empty;
        public string RozvadecCislo { get; set; } = string.Empty;

        /// <summary>Druh zařízení čerpadlo, motor, trafo</summary>
        public string Druh { get; set; } = string.Empty;
        public string Napeti { get; set; } = "400";
        public string Radek { get; set; }

        /// <summary>Rozvaděč</summary>

        // Vypsání hodnot záznamů podle názvů parametrů
        public void Vypis()
        {

            foreach (var Parametr in vlastnosti)
            {
                if (Parametr == "Item") continue;
                //// Pomocí reflexe získáme hodnotu vlastnosti
                //PropertyInfo vlastnost = GetType().GetProperty(Parametr, BindingFlags.Public | BindingFlags.Instance);
                //if (vlastnost != null)
                //{
                //    var hodnota = vlastnost.GetValue(this);
                //    Console.WriteLine($"{Parametr}: {hodnota}");
                //}
                Console.WriteLine($"{Parametr}: {this[Parametr]}");
            }
            Console.WriteLine();
            
        }

        //public static void NastavVlastnost(string nazevVlastnosti, object hodnota)
        //public static void NastavVlastnost(object obj, string nazevVlastnosti, object hodnota)
        //{
        //    // Získáme typ objektu
        //    //var typ = obj.GetType();
        //    //object obj = new Zarizeni();
        //    //var typ = obj.GetType();
        //    // Získáme vlastnost podle názvu
        //    var vlastnost = obj.GetType().GetProperty(nazevVlastnosti);

        //    if (vlastnost != null && vlastnost.CanWrite)
        //    {
        //        // Nastavíme hodnotu vlastnosti
        //        vlastnost.SetValue(obj, Convert.ChangeType(hodnota, vlastnost.PropertyType));
        //    }
        //    else
        //    {
        //        throw new ArgumentException($"Vlastnost {nazevVlastnosti} neexistuje nebo není zapisovatelná.");
        //    }
        //}

        /// <summary>Volání parametru jako string např. Nadpis[Name]  </summary>
        public object this[string nazev]
        {
            get
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance);
                if (prop == null) throw new ArgumentException($"Neexistující vlastnost: {nazev}");
                return prop.GetValue(this);
            }
            set
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance);
                if (prop == null) throw new ArgumentException($"Neexistující vlastnost: {nazev}");
                prop.SetValue(this, Convert.ChangeType(value, prop.PropertyType));
            }
        }

        /// <summary> List vlastností třídy </summary>
        [JsonIgnore]
        public List<string> vlastnosti => GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                      .Select(p => p.Name)
                       .ToList();

    }
}
