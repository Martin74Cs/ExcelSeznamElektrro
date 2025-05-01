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

    //Pokud chceš jednotky ukládat přímo jako atribut, přidej si vlastní:
    [AttributeUsage(AttributeTargets.Property)]
    public class JednotkyAttribute(string text) : Attribute
    {
        public string Text { get; } = text;
    }

    public class Zarizeni : Entity
    {
        //public Zarizeni() { }

        //public Zarizeni(string Tag, string PID, string Popis, string Prikon ,string BalenaJednotka, string Menic , string Nic, string HPstr , string Proud, string PruzezMM2, string AWG, 
        //string Delka, string Delkaft , string Rozvadec , string RozvadecCislo )
        //{
        //    this.Tag = Tag;
        //    this.PID = PID;
        //    this.Popis = Popis;
        //    this.Prikon = Prikon;
        //    this.BalenaJednotka = BalenaJednotka;
        //    this.Menic = Menic;
        //    this.Nic = Nic;
        //    //hp
        //    this.Proud = Proud;
        //    this.PruzezMM2 = PruzezMM2;
        //    this.AWG = AWG;
        //    this.Delka = Delka;
        //    this.Delkaft = Delkaft;
        //    this.Rozvadec = Rozvadec;
        //    this.RozvadecCislo = RozvadecCislo;
        //    //this.Druh = Druh;
        //}
        /// <summary>Druh zařízení čerpadlo, motor, trafo</summary>

        //var TextPole = new string[] { "Tag", "PID", "Equipment name", "kW", "BalenaJednotka", "Menic", "Nic", "Power [HP]", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };

        /// <summary>Označení zařízení </summary>
        public string Tag { get; set; } = string.Empty;
        public string PID { get; set; } = string.Empty;

        /// <summary>Popis zařízení</summary>
        [Display(Name = "Jméno zařízení")]
        public string Popis { get; set; } = string.Empty;

        [Display(Name = "Příkon")]
        [Jednotky("[kW]")]
        public string Prikon { get; set; } = string.Empty;
        public string BalenaJednotka { get; set; } = string.Empty;
        public string Menic { get; set; } = string.Empty;
        public string Nic { get; set; } = string.Empty;

        [JsonIgnore]
        [Display(Name = "Příkon")]
        [Jednotky("[hp]")]
        public double HP => Math.Round(double.TryParse(Prikon, out double hodnota) ? hodnota * 1.341022 : 0, 2); // Převod textu na číslo a na koně

        [Display(Name = "Proud")]
        [Jednotky("[A]")]
        public string Proud { get; set; } = string.Empty;

        [Display(Name = "Průřez")]
        [Jednotky("[mm2]")]
        public string PruzezMM2 { get; set; } = string.Empty;
        public string AWG { get; set; } = string.Empty;

        [Display(Name = "Délka")]
        [Jednotky("[m]")]
        public double Delka { get; set; } = 100;

        [Display(Name = "Délka")]
        [Jednotky("[ft]")]

        public double Delkaft { get; set; } 

        public string Rozvadec { get; set; } = string.Empty;
        public string RozvadecCislo { get; set; } = string.Empty;

        /// <summary>Označení celého rozvaděče</summary>
        [JsonIgnore]
        public string RozvadecOznačení => Rozvadec + " " + RozvadecCislo;

        /// <summary>Druh zařízení čerpadlo, motor, trafo</summary>
        public string Druh { get; set; } = string.Empty;
        public string Napeti { get; set; } = "400";

        /// <summary>Odpovídá radku strojního zařízení</summary>
        public int Radek { get; set; }  
        public string Deleni { get; set; } = string.Empty;

        [JsonIgnore]
        public Kabel Kabel { get; set; } = new();

        [JsonIgnore]
        public Motor Motor { get; set; } = new();

        // Vypsání hodnot záznamů podle názvů parametrů
        /// <summary>Rozvaděč</summary>
        public void Vypis()
        {

            foreach (var Parametr in Vlastnosti)
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
        [JsonIgnore]
        public object? this[string nazev]
        {
            get
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance) ?? throw new ArgumentException($"Neexistující vlastnost: {nazev}");
                return prop.GetValue(this);
            }
            set
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance) ?? throw new ArgumentException($"Neexistující vlastnost: {nazev}");
                prop.SetValue(this, Convert.ChangeType(value, prop.PropertyType));
            }
        }

        /// <summary> List vlastností třídy </summary>
        [JsonIgnore]
        public List<string> Vlastnosti => GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                      .Select(p => p.Name)
                       .ToList();

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
