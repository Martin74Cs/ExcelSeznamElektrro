using Aplikace.Sdilene;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy {
    //public class Pole {
    //    public string BasePath { get; set; } = string.Empty;
    //    public string Místnost { get; set; } = string.Empty;
    //    public string Projekt { get; set; } = string.Empty;
    //    public string Název { get; set; } = string.Empty;
    //    public string Poznámka { get; set; } = string.Empty;
    //    public DateTime Datum { get; set; } = DateTime.Now;
    //}

    //Jedná se o singleton, který uchovává informace o aktuálním projektu a umožňuje jejich načítání a ukládání do souboru v AppData
    public class Informace : IDisposable {

        //skrýtí konstruktoru, aby nebylo možné vytvořit další instance třídy
        private Informace() { }

        private static Informace? Info = null ;
        private static List<KeyValuePair<string, string>> Data = [];
        
        [Display(Name = "Základní složka projektu")]
        public string BasePath { get; set; } = string.Empty;

        [Display(Name = "Základní soubor strojnů")]
        public string SouborStrojeXls { get; set; } = string.Empty;

        [Display(Name = "Stroje.json")]
        public string SouborStrojeJson { get; set; } = string.Empty;

        [Display(Name = "Elektro")]
        public string SouborElektroJson { get; set; } = string.Empty;

        public string Místnost { get; set; } = string.Empty;
        public string Projekt { get; set; } = string.Empty;
        public string Název { get; set; } = string.Empty;
        public string Poznámka { get; set; } = string.Empty;
        public DateTime Datum { get; set; } = DateTime.Now;

        private static readonly string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        private static string Adresar => Path.Combine(appData, "Elektro");
        private static string Soubor => Path.Combine(Adresar, "data.txt");
        public static  Informace Create { 
            get
            {
                if (Info == null)
                {
                    //informace = new Informace();
                    if(File.Exists(Informace.Soubor)) {
                        Nacti();
                    }
                    else
                        Info = new();
                }
                return Info!;
            }

            private set;
        }

        public static void Add(string key, string value) {
            var existing = Data.FirstOrDefault(kv => kv.Key == key);
            if(!existing.Equals(default(KeyValuePair<string, string>))) {
                Data.Remove(existing);
            }
            Data.Add(new KeyValuePair<string, string>(key, value));
        }

        public static string? Get(string key) {
            var existing = Data.FirstOrDefault(kv => kv.Key == key);
            if(!existing.Equals(default(KeyValuePair<string, string>))) {
                return existing.Value;
            }
            return null;
        }

        public void Ulozit() {

            if(!Directory.Exists(Adresar)) {
                Directory.CreateDirectory(Adresar);
            }
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(this, Soubory.Nastaveni());
            File.WriteAllText(Soubor, json);
            //Create = this;
            Console.WriteLine("Cesty k souborům aktualizovány");
        }

        private static void Nacti()
        {
            if (!System.IO.File.Exists(Soubor)) return;
            string jsonString = System.IO.File.ReadAllText(Soubor);
            Info = Newtonsoft.Json.JsonConvert.DeserializeObject<Informace>(jsonString, Soubory.Nastaveni()) ?? new();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Uložení jen při explicitním volání Dispose
                Ulozit();
            }
        }
    }


}
