using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{

    //Pokud chceš jednotky ukládat přímo jako atribut, přidej si vlastní:
    [AttributeUsage(AttributeTargets.Property)]
    public class JednotkyAttribute(string text) : Attribute
    {
        public string Text { get; } = text;
    }

    public class Zarizeni : Entity , INotifyPropertyChanged
    {
        private string tag = string.Empty;
        private string predmet = string.Empty;
        private string prikon = string.Empty;
        private string balenaJednotka = string.Empty;
        private string menic = string.Empty;
        private string nic = string.Empty;
        private string proud = string.Empty;
        private string pruzezMM2 = string.Empty;
        private string pID = string.Empty;
        private int pocet;
        private string popis = string.Empty;
        private string aWG = string.Empty;
        private double delka = 100;
        private double delkaft;
        private string rozvadec = string.Empty;
        private string rozvadecCislo = string.Empty;
        private string druh = string.Empty;
        private Druhy druhenum = Druhy.Rozvadeč; // Výchozí hodnota pro enum Druhy
        private string napeti = "400";
        private int radek;
        private string vodice = string.Empty;
        private Kabel kabel = new();
        private Motor motor = new();
        private string patro = string.Empty;
        private string vykres = string.Empty;
        private bool isExist = false;
        private string bod = string.Empty;
        private bool isExistElektro = false;
        private string bodElektro = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

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
        public string Tag { get => tag; set => SetProperty(ref tag , value); }
        public string Predmet { get => predmet; set => SetProperty(ref predmet , value); }
        public string PID { get => pID; set => SetProperty(ref pID, value); }
        public int Pocet { get => pocet; set => SetProperty(ref pocet, value); }

        /// <summary>Popis zařízení</summary>
        [Display(Name = "Jméno zařízení")]
        public string Popis { get => popis; set => SetProperty(ref popis, value); }
        [Display(Name = "Příkon")]
        [Jednotky("[kW]")]
        public string Prikon { get => prikon; set => SetProperty(ref prikon , value); }
        public string BalenaJednotka { get => balenaJednotka; set => SetProperty(ref balenaJednotka, value); }
        public string Menic { get => menic; set => SetProperty(ref menic, value); }
        public string Nic { get => nic; set => SetProperty(ref nic, value); }

        [JsonIgnore]
        [Display(Name = "Příkon")]
        [Jednotky("[hp]")]
        public double HP => Math.Round(double.TryParse(Prikon, out double hodnota) ? hodnota * 1.341022 : 0, 2); // Převod textu na číslo a na koně

        [Display(Name = "Proud")]
        [Jednotky("[A]")]
        public string Proud { get => proud; set => SetProperty(ref proud, value); }

        [Display(Name = "Průřez")]
        [Jednotky("[mm2]")]
        public string PruzezMM2 { get => pruzezMM2; set => SetProperty(ref pruzezMM2, value); }

        public string AWG { get => aWG; set => SetProperty(ref aWG, value); }

        [Display(Name = "Délka")]
        [Jednotky("[m]")]
        public double Delka { get => delka; set => SetProperty(ref delka, value); }
        [Display(Name = "Délka")]
        [Jednotky("[ft]")]

        public double Delkaft { get => delkaft; set => SetProperty(ref delkaft, value); }

        public string Rozvadec { get => rozvadec; set => SetProperty(ref rozvadec, value); }
        public string RozvadecCislo { get => rozvadecCislo; set => SetProperty(ref rozvadecCislo, value); }
        /// <summary>Označení celého rozvaděče</summary>
        [JsonIgnore]
        public string RozvadecOznačení => Rozvadec + " " + RozvadecCislo;
        public string Vyvod { get; set; } = string.Empty;
        /// <summary>Druh zařízení čerpadlo, motor, trafo</summary>
        public string Druh { get => druh; set => SetProperty(ref druh, value); }

        [JsonIgnore]
        public Druhy DruhEnum { get => druhenum; set => SetProperty(ref druhenum, value); }
        public string Napeti { get => napeti; set => SetProperty(ref napeti, value); }
        /// <summary>Odpovídá radku strojního zařízení</summary>
        public int Radek { get => radek; set => SetProperty(ref radek, value); }

        /// <summary>Odpovídá počtu silových vodičů </summary>
        public string Vodice { get => vodice; set => SetProperty(ref vodice, value); }

        //[JsonIgnore]
        public Kabel Kabel { get => kabel; set => SetProperty(ref kabel, value); }
        //[JsonIgnore]
        public Motor Motor { get => motor; set => SetProperty(ref motor, value); }
        public string Patro { get => patro; set => SetProperty(ref patro, value); }
        public string Vykres { get => vykres; set => SetProperty(ref vykres, value); }         /// <summary>false=neexistuje</summary>
        public bool IsExist { get => isExist; set => SetProperty(ref isExist, value); }

        [JsonConverter(typeof(PointToStringConverter))]
        public string Bod { get => bod; set => SetProperty(ref bod, value); }

        /// <summary>Definice bloku elektro</summary>
        public bool IsExistElektro { get => isExistElektro; set => SetProperty(ref isExistElektro, value); }

        public double Otoceni { get; set; } = 0.0;

        [JsonConverter(typeof(PointToStringConverter))]
        public string BodElektro { get;  set;} = string.Empty; // = MyPoint3d.Origin;

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
        public List<string> Vlastnosti => [.. GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance).Select(p => p.Name)];

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
        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            return true;
        }

        public enum DruhZarizeni
        {
            Motor,
            Cerpadlo,
            Trafo,
            Jine
        }

        public enum Druhy
        {
            Přívod,
            Spojka,
            Rozvadeč,
            Motor,
            Nic,
        }
    }

    public class MyPoint3d(double x, double y, double z)
    {
        public double X { get; set; } = x;
        public double Y { get; set; } = y;
        public double Z { get; set; } = z;

        public static MyPoint3d Origin => new(0.0, 0.0, 0.0);
    }

    public class PointToStringConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            // Převádíme JSON objekt na string v modelu
            return objectType == typeof(string);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, Newtonsoft.Json.JsonSerializer serializer)
        {
            var token = JToken.Load(reader);

            if (token.Type == JTokenType.Object)
            {
                var x = token["X"]?.Value<double>() ?? 0.0;
                var y = token["Y"]?.Value<double>() ?? 0.0;
                var z = token["Z"]?.Value<double>() ?? 0.0;

                return $"X={x:0.00},Y={y:0.00},Z={z:0.00}";
            }

            return token.ToString();
        }

        public override void WriteJson(JsonWriter writer, object value, Newtonsoft.Json.JsonSerializer serializer)
        {
            var text = value?.ToString() ?? "";

            // Pokus o parsování zpět na objekt, pokud se bude serializovat zpět
            var parts = text.Split(',');
            try
            {
                var obj = new JObject();
                foreach (var part in parts)
                {
                    var kv = part.Split('=');
                    if (kv.Length == 2)
                        obj[kv[0].Trim()] = double.Parse(kv[1]);
                }

                obj.WriteTo(writer);
            }
            catch
            {
                writer.WriteValue(value);
            }
        }
    }

}


