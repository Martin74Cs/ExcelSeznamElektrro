using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using static Aplikace.Tridy.Motor;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Upravy
{
    public static class LigthChem
    {
        /// <summary>Vytvořit z Seznamu strojů json a Csv pro další doplnění</summary>
        public static void StrojniToJsonCsv()
        {
            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            
            if (!Directory.Exists(Cesty.Elektro))
                Directory.CreateDirectory(Cesty.Elektro);

            string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            var Stara = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Stara.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));

            string cestaData = Path.Combine(Cesty.ElektroDataCsv);
            if (!File.Exists(cestaData))
            {
                //vytvoření základu pro json jen pokud neexistuje.
                Stara.SaveJsonList(Path.ChangeExtension(cestaData, ".json"));
                //vytvoření csv pro doplnění
                Stara.SaveToCsv(cestaData);
            }
            else
            {
                Console.WriteLine($"Vyvoření kopie {Path.GetFileName(cestaData)} přeskočeno");
            }

        }

        /// <summary>Vytvoření excelu dle ElektroRozvaděč.Json</summary>
        public static void JsonToExcel()
        {
            string cestaData = Cesty.ElektroDataJson;
            var Stara = Soubory.LoadJsonList<Zarizeni>(cestaData);

            //Možná proud asi jen tam kde není.
            //Stara.AddProud();

            //Přidání typu kabelu.
            Stara.AddKabelCyky(1.5);

            string filename = "Seznam.xlsx";
            var cesta = Path.Combine(Cesty.Elektro, filename);
            //vytvoření nebo otevření dokumentu elektro
            var ExcelApp = new ExcelApp(cesta);

            //Nastavení nebo vytvoření záložky
            ExcelApp.GetSheet("Seznam Elektro");
            //Vytvoření nadpisů
            var range = ExcelApp.Nadpisy([.. Nadpis.DataCz()]);

            //Formátování nadpisů
            ExcelApp.NadpisSet(range);

            //čísla sloupců a nazvy tříd. 
            var dir = new Dictionary<int, string>() {
                {1,"Tag"},
                {2,"Popis"},
                {3,"Prikon"},
                {4,"Napeti"},
                {5,"Proud"},
                {6,"BalenaJednotka"},
                {7,"Menic"},
                {8,"Druh"},
                {9,"PruzezMM2"},
                {10,"Delka"},
                {11,"Rozvadec"},
                {12,"RozvadecCislo"},

                //{100,"PID"},
                //{101,"Nic"},
                //{102,"HP"},
                //{103,"AWG"},
                //{105,"Napeti"},
                //{106,"Radek"},
            };
            ExcelApp.ClassToExcel(Row: 3, Stara, dir);

            Console.WriteLine("Probíhá načítaní kabelů");
            //Vytvoření pole kabelů pro zápis do Excelu
            var PoleData = KabelList.Kabely(Stara);

            //Nová záložka nebo nastav existující
            ExcelApp.GetSheet("Kabely");

            //Doplnení nadpisu a ramecku
            ExcelApp.ExcelSaveNadpis(PoleData);

            //do Excel vyplní od radku 3 data data z PoleData mělo by se jednat o seznam kabelů
            ExcelApp.KabelyToExcel(PoleData, 3);

            //Vyzváření seznamu kabelů podle krytérii
            Pridat.Soucet(ExcelApp, PoleData);
             
            ExcelApp.ExcelQuit(cesta);

        }

        public static void Hlavni()
        {
            string cesta = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string json = Path.ChangeExtension(cesta, ".json");
            if (!File.Exists(json))
                return;

            //var Source = Soubory.LoadFromCsv<Zarizeni>(cesta);
            //Prevod.UpdateCsvToJson(Source, Target);
        }

        public static void AddProud()
        {
            string cestaData = Cesty.ElektroDataJson;
            var Stara = Soubory.LoadJsonList<Zarizeni>(cestaData);

            //Možná proud asi jen tam kde není.
            var Add = Stara.AddProud();
            Add.SaveJsonList(cestaData);
        }



        /// <summary>Seznam vývodů pro doplnění </summary>
        public static void AddVyvody()
        {
            //Aktualní seznam vývodů
            string cesta = Cesty.ElektroDataJson;
            var Target = Soubory.LoadJsonList<Zarizeni>(cesta);

            //Seznam vývodů pro doplnění
            var Add = new List<Zarizeni>();
            string cesta1 = Path.Combine(Cesty.Elektro, @"Vývody.csv");
            if (!File.Exists(cesta1))
            {
                //Pokud neexistuje tak vytvoř
                Add.Add(new Zarizeni());
                Add.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
                Add.SaveToCsv(cesta1);
                return;
            }

            //Načtení vývodů pro doplnnění 
            Add = Soubory.LoadFromCsv<Zarizeni>(cesta1);
            bool zmena = false;
            for (int i = 0; i < Add.Count; i++)
            {
                //přidat identifikátor
                if (string.IsNullOrEmpty(Add[i].Apid))
                { 
                    Add[i].Apid = ExcelLoad.Apid();
                    zmena = true;
                }
                //
            }

            //Mělo by přídat proud
            Add.AddProud();

            //pokud nexistuje Apid tak přepsat Json
            if (zmena)
            { 
                //Prevod.SaveToCsv(Add, cesta1);
                //Add.SaveToCsv(cesta1);
                Add.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            }

            int pocet = Target.Count + 1;
            foreach (var item in Add)
            {
                //hledej existeci podle Apid
                var Toto = Target.FirstOrDefault(x => x.Apid == item.Apid);

                //Existuje tak ho smaž
                if (Toto != null) //continue;
                    Target.Remove(Toto);

                //znovu přidat
                item.Radek = pocet++;
                Target.Add(item);
            }

            //uložit upravená data do Json
            Target.SaveJsonList(cesta);

            //Target.SaveToCsv(Path.ChangeExtension(cesta, ".csv"));
            //Target.SaveJsonList(Path.ChangeExtension(cesta, ".txt"));
        }

        public static void AddKabely()
        {
            //string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string cesta1 = Path.Combine(Cesty.ElektroDataJson);
            var Target = Soubory.LoadJsonList<Zarizeni>(cesta1);

            Target.AddKabelCyky(1.5);

            Target.SaveJsonList(cesta1);
            //Target.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));
        }

        public static List<Zarizeni> AddKabelCyky(this List<Zarizeni> Target, double rezerva = 1.5)
        {
            var KabelCu = Soubory.LoadJsonListEn<KabelVse>(Cesty.CuJson)
                .Where(x => x.Name.Contains("CYKY", StringComparison.OrdinalIgnoreCase) && x.Deleni == "4")
                .OrderBy(x => x.IzAE).ToList();

            //vymazaní položek ze seznam
            KabelCu.RemoveAll(x => x.SLmm2 < 2.5);

            //Kontrola
            //KabelCu.SaveJsonList(Path.ChangeExtension(cesta1, ".txt"));

            var KabelAL = Soubory.LoadJsonList<KabelVse>(Cesty.AlJson).OrderByDescending(x => x.IzAE ).ToList(); 
            //vymazaní položek ze seznam
            KabelAL.RemoveAll(x => x.SLmm2 < 16);

            //item nevím co to je
            var properties = typeof(Zarizeni).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => p.CanWrite && p.Name != "Item")
                                 .ToList();

            for (int i = 0; i < Target.Count; i++)
            {
                foreach (var prop in properties)
                {
                    //var value = prop.GetValue(target);
                    var proud = double.TryParse(Target[i].Proud, out var p) ? p : 1;
                    var JedenKabel = KabelCu.FirstOrDefault(x => x.MaxProud > proud * rezerva);
                    if (JedenKabel == null) continue;

                    //prop.SetValue(target, value);
                    //target.GetType().GetProperty("PruzezMM2").SetValue(target, JedenKabel.SLmm2.Tostring());
                    Target[i].PruzezMM2 = JedenKabel.SLmm2.ToString();
                    Target[i].Deleni = JedenKabel.Deleni;
                    Target[i].Kabel = JedenKabel;
                }
            }
            return Target;
        }

        public static void DoplneniCsvToJson()
        {
            //Soubor kam bude doplněno
            string cestaData = Path.Combine(Cesty.ElektroDataJson);
            var Target = Soubory.LoadJsonList<Zarizeni>(cestaData);

            //Data pro doplnění
            string cesta = Path.Combine(Cesty.ElektroDataCsv);
            if (!File.Exists(cesta)) return;
            var Source = Soubory.LoadFromCsv<Zarizeni>(cesta);
            
            //Doplnění dat do Json.
            Prevod.UpdateCsvToJson(Source, Target );

            //uložit doplnění informace do Json
            Target.SaveJsonList(cestaData);
        }

        public static void VyvoritFMKM()
        {
            //var cesta = Environment.ProcessPath;            // Získá úplnou cestu ke spuštěnému procesu
            //var dir = Path.GetDirectoryName(cesta);         // Získá adresář, kde je spustitelný soubor
            VyvoritFM();
            VyvoritKM();
        }

        /// <summary> Převod seznamu frekvenčních měničů na Json </summary>
        public static void VyvoritFM()
        {
            string basePath = Path.Combine(Cesty.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv"); 
           
            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary> Převod seznamu stykačů na Json </summary>
        public static void VyvoritKM()
        {
            string basePath = Path.Combine(Cesty.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv"); 
           
            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary> Převod seznamu motorů a motorů3000 na jeden Json </summary>
        public static void VyvoritMotor()
        {
            string basePath = Path.Combine(Cesty.Data, "Motory");
            string CestaMotor = Path.Combine(basePath, "Motory.csv");
            var Motor = Soubory.LoadFromCsv<Motor>(CestaMotor);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            string CestaMotor3000 = Path.Combine(basePath, "Motory3000.csv");
            var Motor3000 = Soubory.LoadFromCsv<Motor>(CestaMotor3000);
            Console.WriteLine($"Pocet motorů: {Motor3000.Count}");

            Motor.AddRange(Motor3000);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            Motor.SaveJsonList(Path.ChangeExtension(CestaMotor, ".json"));
            Console.WriteLine($"Motory uloženy jako Json");
        }
    }
}
