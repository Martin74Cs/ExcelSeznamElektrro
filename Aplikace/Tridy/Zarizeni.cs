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
        public event PropertyChangedEventHandler? PropertyChanged;

        #region 1. Identifikace a popis
        private string tag = string.Empty;
        [Category("1. Identifikace a popis")]
        [DisplayName("Označení (Tag)")]
        [Description("Unikátní kód / označení zařízení.")]
        public string Tag { get => tag; set => SetProperty(ref tag , value); }

        private string tagStroj = string.Empty;
        [Category("1. Identifikace a popis")]
        [DisplayName("Tag stroje")]
        [Description("Označení odpovídajícího strojního zařízení.")]
        public string TagStroj { get => tagStroj; set => SetProperty(ref tagStroj , value); }

        private string predmet = string.Empty;
        [Category("1. Identifikace a popis")]
        [DisplayName("Předmět")]
        [Description("Předmět / název zařízení.")]
        public string Predmet { get => predmet; set => SetProperty(ref predmet , value); }

        private string pid = string.Empty;
        [Category("1. Identifikace a popis")]
        [DisplayName("P&ID Schéma")]
        [Description("Označení příslušného P&ID technologického schématu.")]
        [Display(Name = "Technologické schéma")]
        public string Pid { get => pid; set => SetProperty(ref pid, value); }

        private string popis = string.Empty;
        [Category("1. Identifikace a popis")]
        [DisplayName("Popis zařízení")]
        [Description("Jméno nebo popis zařízení.")]
        [Display(Name = "Jméno zařízení")]
        public string Popis { get => popis; set => SetProperty(ref popis, value); }

        private string pozice = string.Empty;
        [Category("1. Identifikace a popis")]
        [DisplayName("Pozice")]
        [Description("Označení pozice.")]
        public string Pozice { get => pozice; set => SetProperty(ref pozice, value); }

        private string poznamka = string.Empty;
        [Category("1. Identifikace a popis")]
        [DisplayName("Poznámka")]
        [Description("Doplňující poznámky k zařízení.")]
        public string Poznamka { get => poznamka; set => SetProperty(ref poznamka, value); }

        private int pocet;
        [Category("1. Identifikace a popis")]
        [DisplayName("Počet")]
        [Description("Počet kusů.")]
        public int Pocet { get => pocet; set => SetProperty(ref pocet, value); }

        private int radek;
        [Category("1. Identifikace a popis")]
        [DisplayName("Řádek stroje")]
        [Description("Řádek odpovídající strojnímu zařízení.")]
        public int Radek { get => radek; set => SetProperty(ref radek, value); }
        #endregion

        #region 2. Umístění a fáze
        private string patro = string.Empty;
        [Category("2. Umístění a fáze")]
        [DisplayName("Patro")]
        [Description("Patro budovy / podlaží.")]
        public string Patro { get => patro; set => SetProperty(ref patro, value); }

        private string vykres = string.Empty;
        [Category("2. Umístění a fáze")]
        [DisplayName("Výkres")]
        [Description("Identifikační číslo výkresu.")]
        public string Vykres { get => vykres; set => SetProperty(ref vykres, value); }

        private string etapa = string.Empty;
        [Category("2. Umístění a fáze")]
        [DisplayName("Fáze výstavby")]
        [Description("Fáze nebo etapa výstavby.")]
        [Display(Name = "Fáze výstavby")]
        public string Etapa { get => etapa; set => SetProperty(ref etapa, value); }
        #endregion

        #region 3. Elektrické parametry
        private string prikon = string.Empty;
        [Category("3. Elektro parametry")]
        [DisplayName("Příkon elektro")]
        [Description("Elektrický příkon zařízení [kW].")]
        [Display(Name = "Příkon elektro")]
        [Jednotky("[kW]")]
        public string Prikon { get => prikon; set => SetProperty(ref prikon , value); }

        private string prikonStroj = string.Empty;
        [Category("3. Elektro parametry")]
        [DisplayName("Příkon strojní")]
        [Description("Strojní příkon zařízení [kW].")]
        [Display(Name = "Příkon strojní")]
        [Jednotky("[kW]")]
        public string PrikonStroj { get => prikonStroj; set => SetProperty(ref prikonStroj, value); }

        private string proud = string.Empty;
        [Category("3. Elektro parametry")]
        [DisplayName("Proud")]
        [Description("Jmenovitý proud [A].")]
        [Display(Name = "Proud")]
        [Jednotky("[A]")]
        public string Proud { get => proud; set => SetProperty(ref proud, value); }

        private string napeti = "400";
        [Category("3. Elektro parametry")]
        [DisplayName("Napětí")]
        [Description("Jmenovité napětí [V].")]
        [Display(Name = "Napětí")]
        public string Napeti { get => napeti; set => SetProperty(ref napeti, value); }

        [JsonIgnore]
        [Category("3. Elektro parametry")]
        [DisplayName("Příkon (HP)")]
        [Description("Přepočtený příkon v koních (pouze pro čtení).")]
        [Display(Name = "Příkon")]
        [Jednotky("[hp]")]
        public double HP => Math.Round(double.TryParse(Prikon, out double hodnota) ? hodnota * 1.341022 : 0, 2); // Převod textu na číslo a na koně
        #endregion

        #region 4. Napájení a řízení
        private string druh = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Druh zařízení")]
        [Description("Druh zařízení (Motor, Přívod, Spojka, Rozvaděč atd.).")]
        [Display(Name = "Druh zařízení")]
        public string Druh { get => druh; set => SetProperty(ref druh, value); }

        private string typ = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Typ")]
        [Description("Konkrétní typ zařízení (např. čerpadlo, vývěva, míchadlo).")]
        public string Typ { get => typ; set => SetProperty(ref typ, value); }

        private Druhy druhenum = Druhy.Rozvadeč; // Výchozí hodnota pro enum Druhy
        [JsonIgnore]
        [Category("4. Napájení a řízení")]
        [DisplayName("Druh (Enum)")]
        [Description("Výběr druhu ze seznamu možností (Enum).")]
        public Druhy DruhEnum { get => druhenum; set => SetProperty(ref druhenum, value); }

        private string menic = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Měnič")]
        [Description("Typ nebo přítomnost frekvenčního měniče.")]
        [Display(Name = "Měnič")]
        public string Menic { get => menic; set => SetProperty(ref menic, value); }

        private string balenaJednotka = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Balená jednotka")]
        [Description("Označení balené jednotky (Package Unit).")]
        public string BalenaJednotka { get => balenaJednotka; set => SetProperty(ref balenaJednotka, value); }

        private string nic = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Nic (Rezerva)")]
        [Description("Rezervní pole.")]
        public string Nic { get => nic; set => SetProperty(ref nic, value); }

        [Category("4. Napájení a řízení")]
        [DisplayName("Vývod")]
        [Description("Označení vývodu z rozvaděče.")]
        public string Vyvod { get; set; } = string.Empty;

        private string vodice = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Počet silových vodičů")]
        [Description("Odpovídá počtu silových vodičů.")]
        public string Vodice { get => vodice; set => SetProperty(ref vodice, value); }

        private string rozvadec = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Rozvaděč (Text)")]
        [Description("Označení nebo zkratka rozvaděče.")]
        [Display(Name = "Rozvaděč text")]
        public string Rozvadec { get => rozvadec; set => SetProperty(ref rozvadec, value); }

        private string rozvadecCislo = string.Empty;
        [Category("4. Napájení a řízení")]
        [DisplayName("Rozvaděč (Číslo)")]
        [Description("Doplňující číslo rozvaděče.")]
        [Display(Name = "Rozvaděč číslo")]
        public string RozvadecCislo { get => rozvadecCislo; set => SetProperty(ref rozvadecCislo, value); }

        [JsonIgnore]
        [Category("4. Napájení a řízení")]
        [DisplayName("Označení rozvaděče")]
        [Description("Celé spojené označení rozvaděče (pouze pro čtení).")]
        [Display(Name = "Označení rozvaděče")]
        public string RozvadecOznačení => Rozvadec + RozvadecCislo;
        #endregion

        #region 5. Kabelové připojení
        private Kabel kabel = new();
        [Category("5. Kabelové připojení")]
        [DisplayName("Kabel (Objekt)")]
        [Description("Interní data kabelu.")]
        public Kabel Kabel { get => kabel; set => SetProperty(ref kabel, value); }

        private int pocetKabelu = 1; //string.Empty;
        [Category("5. Kabelové připojení")]
        [DisplayName("Počet kabelů")]
        [Description("Počet kabelů [ks].")]
        [Display(Name = "Počet kabelů")]
        [Jednotky("[ks]")]
        public int PocetKabelu { get => pocetKabelu; set => SetProperty(ref pocetKabelu, value); }

        private string prurezMM2 = string.Empty;
        [Category("5. Kabelové připojení")]
        [DisplayName("Průřez [mm2]")]
        [Description("Průřez silových vodičů v mm2.")]
        [Display(Name = "Průřez")]
        [Jednotky("[mm2]")]
        public string PrurezMM2 { get => prurezMM2; set => SetProperty(ref prurezMM2, value); }

        private string aWG = string.Empty;
        [Category("5. Kabelové připojení")]
        [DisplayName("AWG")]
        [Description("Průřez kabelu v jednotkách AWG.")]
        public string AWG { get => aWG; set => SetProperty(ref aWG, value); }

        private double delka = 100;
        [Category("5. Kabelové připojení")]
        [DisplayName("Délka [m]")]
        [Description("Délka kabelové trasy v metrech.")]
        [Display(Name = "Délka")]
        [Jednotky("[m]")]
        public double Delka { get => delka; set => SetProperty(ref delka, value); }

        private double delkaft;
        [Category("5. Kabelové připojení")]
        [DisplayName("Délka [ft]")]
        [Description("Délka kabelové trasy ve stopách.")]
        [Display(Name = "Délka")]
        [Jednotky("[ft]")]
        public double Delkaft { get => delkaft; set => SetProperty(ref delkaft, value); }
        #endregion

        #region 6. Motor
        private Motor motor = new();
        [Category("6. Motor")]
        [DisplayName("Motor (Objekt)")]
        [Description("Interní data motoru.")]
        public Motor Motor { get => motor; set => SetProperty(ref motor, value); }
        #endregion

        #region 7. CAD koordinace
        private bool isExist = false;
        [Category("7. CAD koordinace")]
        [DisplayName("Existuje v projektu")]
        [Description("Indikuje, zda zařízení fyzicky existuje (false = neexistuje).")]
        public bool IsExist { get => isExist; set => SetProperty(ref isExist, value); }

        private string bod = string.Empty;
        [Category("7. CAD koordinace")]
        [DisplayName("Bod (Souřadnice)")]
        [Description("Souřadnice bodu v CAD.")]
        [JsonConverter(typeof(PointToStringConverter))]
        public string Bod { get => bod; set => SetProperty(ref bod, value); }

        private bool isExistElektro = false;
        [Category("7. CAD koordinace")]
        [DisplayName("Elektro blok existuje")]
        [Description("Indikuje přítomnost elektro bloku v CAD.")]
        public bool IsExistElektro { get => isExistElektro; set => SetProperty(ref isExistElektro, value); }

        [Category("7. CAD koordinace")]
        [DisplayName("Otočení bloku")]
        [Description("Úhel otočení elektro bloku v CAD.")]
        public double Otoceni { get; set; } = 0.0;

        [Category("7. CAD koordinace")]
        [DisplayName("Bod Elektro")]
        [Description("Souřadnice elektro bodu v CAD.")]
        [JsonConverter(typeof(PointToStringConverter))]
        public string BodElektro { get;  set;} = string.Empty; // = MyPoint3d.Origin;
        #endregion

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
            Otop,
            Trafo,
            Nic,
        }

        public static T Clone<T>(T source) where T : new()
        {
            if (source == null) return default;

            T copy = new();
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.CanWrite)
                .Where(p => !p.PropertyType.IsEnum) // vyloučí enumy
                .Where(p => p.Name != "Vlastnosti") // vyloučí konkrétní název
                .Where(p => !Attribute.IsDefined(p, typeof(JsonIgnoreAttribute))); // vyloučí [JsonIgnore]

            foreach (var prop in properties)
            {
                var value = prop.GetValue(source);
                prop.SetValue(copy, value);
            }

            return copy;
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


