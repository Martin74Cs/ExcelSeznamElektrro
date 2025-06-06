using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using static Aplikace.Tridy.Motor;
using static Aplikace.Tridy.Zarizeni;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Upravy
{
    public static class LigthChem
    {
        /// <summary>Vytvořit z Seznamu strojů json a Csv pro další doplnění</summary>
        public static void StrojniToJsonCsv()
        {
            //Převod->json,csv
            
            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            if (!Directory.Exists(Cesty.Elektro))
                Directory.CreateDirectory(Cesty.Elektro);

            string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            var Stara = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Stara.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));

            //string cestaData = Path.Combine(Cesty.ElektroDataCsv);
            //if (!File.Exists(cestaData))
            //{
            //    //vytvoření základu pro json jen pokud neexistuje.
            //    Stara.SaveJsonList(Path.ChangeExtension(cestaData, ".json"));
            //    //vytvoření csv pro doplnění
            //    Stara.SaveToCsv(cestaData);
            //}
            //else
            //{
            //    Console.WriteLine($"Vyvoření kopie {Path.GetFileName(cestaData)} přeskočeno");
            //}
        }

        /// <summary>Převod extrahovaných dat z Dwg do Xls s následným převodem do Json</summary>
        public static void DwgXlsToJsonCsv()
        { 
            if (!Directory.Exists(Cesty.Elektro))
                Directory.CreateDirectory(Cesty.Elektro);

            string cesta1 = Path.Combine(Cesty.Elektro, "Pid", @"UpravaZnovu.006.xlsm");
            var Stara = ExcelLoad.DwgDataExcel(cesta1, "Summary", 3);

            Console.WriteLine($"Načteno {Stara.Count} záznamů z {cesta1}");

            //Převod->json,csv 
            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Stara.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));
        }


        /// <summary>Vytvoření excelu dle ElektroRozvaděč.Json</summary>
        public static void JsonToExcel()
        {
            string cestaData = Cesty.ElektroDataJson;
            var Stara = Soubory.LoadJsonList<Zarizeni>(cestaData);

            //Vývody pro doplnění
            var cestaVývody = Path.Combine(Cesty.Elektro, "Vývody.json");
            var Vývody = Soubory.LoadJsonList<Zarizeni>(cestaVývody);

            //Spojení původních s doplněnými
            Stara.Concat(Vývody);

            //Možná proud asi jen tam kde není.
            //Stara.AddProud();

            //Přidání typu kabelu pokud chybí .
            var prazdne = Stara.Where(x => x.Kabel == null).ToList();
            prazdne.AddKabelCyky(1.6);

            // Spojení původních neprázdných s doplněnými kabely – Concat vytvoří novou spojenou kolekci
            //Stara = Stara.Where(x => x.Kabel != null).Concat(prazdne).ToList();
            Stara = [.. Stara.Where(x => x.Kabel != null), .. prazdne];

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
                {13,"Predmet"},

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
            var DataTrida = KabelList.KabelyTrida(Stara);

            var Change = new List<Trasa>();
            foreach (var kabel in DataTrida)
            {
                if(kabel.Hlavni != null) 
                    Change.Add(kabel.Hlavni);
                if(kabel.PTC != null)
                    Change.Add(kabel.PTC);
                if(kabel.Ovladani != null)
                    Change.Add(kabel.Ovladani);
            }

            //pole kabelů zapsat do Excel tabulky 
            var PoleData = KabelList.KabelyTridaToString(Change);
            //var PoleData = KabelList.Kabely(Stara);

            //Nová záložka nebo nastav existující
            ExcelApp.GetSheet("Kabely");

            //Doplnení nadpisu a ramecku
            ExcelApp.ExcelSaveNadpis(PoleData);

            //Do Excel vyplní od radku 3 data z PoleData mělo by se jednat o seznam kabelů
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
            ////Aktualní seznam vývodů
            //string cesta = Cesty.ElektroDataJson;
            //var Target = Soubory.LoadJsonList<Zarizeni>(cesta);

            ////Seznam vývodů pro doplnění
            //var Add = new List<Zarizeni>();
            //string cesta1 = Path.Combine(Cesty.Elektro, @"Vývody.csv");
            //if (!File.Exists(cesta1))
            //{
            //    //Pokud neexistuje tak vytvoř
            //    Add.Add(new Zarizeni());
            //    Add.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            //    Add.SaveToCsv(cesta1);
            //    return;
            //}

            ////Načtení vývodů pro doplnnění 
            //Add = Soubory.LoadFromCsv<Zarizeni>(cesta1);
            //bool zmena = false;
            //for (int i = 0; i < Add.Count; i++)
            //{
            //    //přidat identifikátor
            //    if (string.IsNullOrEmpty(Add[i].Apid))
            //    { 
            //        Add[i].Apid = ExcelLoad.Apid();
            //        zmena = true;
            //    }
            //    //
            //}

            ////Mělo by přídat proud
            //Add.AddProud();

            ////pokud nexistuje Apid tak přepsat Json
            //if (zmena)
            //{ 
            //    //Prevod.SaveToCsv(Add, cesta1);
            //    //Add.SaveToCsv(cesta1);
            //    Add.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            //}

            //int pocet = Target.Count + 1;
            //foreach (var item in Add)
            //{
            //    //hledej existeci podle Apid
            //    var Toto = Target.FirstOrDefault(x => x.Apid == item.Apid);

            //    //Existuje tak ho smaž
            //    if (Toto != null) //continue;
            //        Target.Remove(Toto);

            //    //znovu přidat
            //    item.Radek = pocet++;
            //    Target.Add(item);
            //}

            ////uložit upravená data do Json
            //Target.SaveJsonList(cesta);

            ////Target.SaveToCsv(Path.ChangeExtension(cesta, ".csv"));
            ////Target.SaveJsonList(Path.ChangeExtension(cesta, ".txt"));
        }

        //public static void AddKabely()
        //{
        //    //string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
        //    string cesta1 = Path.Combine(Cesty.ElektroDataJson);
        //    var Target = Soubory.LoadJsonList<Zarizeni>(cesta1);

        //    Target.AddKabelCyky(1.6);

        //    Target.SaveJsonList(cesta1);
        //    //Target.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));
        //}

        public static List<Zarizeni> AddKabelCyky(this List<Zarizeni> Target, double rezerva = 1.5)
        {
            var KabelCu = Soubory.LoadJsonListEn<KabelVse>(Cesty.CuJson)
                .Where(x => x.Name.Contains("CYKY", StringComparison.OrdinalIgnoreCase) && x.Deleni == "4")
                .OrderBy(x => x.IzAE).ToList();
            //vymazaní položek ze seznam
            KabelCu.RemoveAll(x => x.SLmm2 < 2.5);

            if(KabelCu.Count < 1) Console.WriteLine("Chyba hledání kabelů Cu");
            else Console.WriteLine($"Kabelů Cu je {KabelCu.Count}");

            //Kontrola
            //KabelCu.SaveJsonList(Path.ChangeExtension(cesta1, ".txt"));

            var KabelAL = Soubory.LoadJsonList<KabelVse>(Cesty.AlJson).OrderByDescending(x => x.IzAE ).ToList(); 
            //vymazaní položek ze seznam
            KabelAL.RemoveAll(x => x.SLmm2 < 16);
            if(KabelAL.Count < 1) Console.WriteLine("Chyba hledání kabelů Al");
            else Console.WriteLine($"Kabelů Al je {KabelAL.Count}");

            var properties = typeof(Zarizeni).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => p.CanWrite && p.Name != "Item")
                                 .ToList();

            for(int i = 0; i < Target.Count; i++)
            {
                int pocet = 1;
                bool volba = true;
                while (volba)
                {
                    foreach (var prop in properties)
                    {
                        //var value = prop.GetValue(target);
                        var proud = double.TryParse(Target[i].Proud, out var p) ? p : 1;
                        //var JedenKabel = KabelCu.FirstOrDefault(x => x.MaxProudVzduch > proud * rezerva) ?? KabelCu.FirstOrDefault(x => pocet * x.MaxProud > proud * rezerva);
                        var JedenKabel = KabelCu.FirstOrDefault(x => x.MaxProudVzduch > proud * rezerva);
                        JedenKabel ??= KabelCu.FirstOrDefault(x => pocet * x.MaxProud > proud * rezerva);
                        if (JedenKabel == null) continue;

                        //prop.SetValue(target, value);
                        //target.GetType().GetProperty("PruzezMM2").SetValue(target, JedenKabel.SLmm2.Tostring());
                        if(pocet > 1) Target[i].PruzezMM2 = pocet.ToString() + "x"+ JedenKabel.SLmm2.ToString();
                        else Target[i].PruzezMM2 = JedenKabel.SLmm2.ToString();
                        Target[i].Vodice = JedenKabel.Deleni;
                        Target[i].Kabel = JedenKabel;
                        volba = false;
                        break;
                    }
                    pocet++;
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

        public static void Rozvadec() {
            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            var Vývody = Path.Combine(Cesty.Elektro, "Vývody.json");
            var Data2 = Soubory.LoadJsonList<Zarizeni>(Vývody);

            Data = [.. Data, .. Data2];

            //rozdělení podle rozvaděče
            var skupiny = Data.GroupBy(x => x.RozvadecOznačení);
            int PocetRozvadecu = skupiny.Count();
            Console.WriteLine($"\nPočet rozvaděčů : {PocetRozvadecu}");

            foreach (var skupina in skupiny) {
                //prvni položka skupiny
                var Jedna = skupina.ElementAtOrDefault(1) ?? new Zarizeni();

                // Převod stringu na enum
                StringToEnum(skupina);

                //Srovnání podle enumu
                var Pole = skupina.OrderByDescending(x => double.TryParse(x.Prikon, NumberStyles.Any, CultureInfo.InvariantCulture, out double result) ? result : 0.0)
                    .OrderBy(x => x.DruhEnum).ToList();

                //Součet příkonů
                var SumaPrikon = Pole.Where(x => x.DruhEnum != Druhy.Přívod && x.DruhEnum != Druhy.Spojka)
                    .Sum(x => double.TryParse(x.Prikon, NumberStyles.Any, CultureInfo.InvariantCulture, out double result) ? result : 0.0);
                Console.WriteLine($" ");
                Console.WriteLine($"\nRozvaděč: {skupina.Key}");
                foreach(var item in Pole) {
                    //Console.WriteLine($"Rozvaděč: {skupina.Key}, Tag: {item.Tag}, Popis: {item.Popis}");
                    if(item.DruhEnum == Druhy.Přívod)
                        Console.WriteLine($"Tag: {item.Tag.Replace("\n"," "),-12}, Druh: {item.Druh,-12}, Popis: {item.Popis,-35}, SumaPříkon: {SumaPrikon,-15:F2}");
                    else
                        Console.WriteLine($"Tag: {item.Tag.Replace("\n", " "),-12}, Druh: {item.Druh,-12}, Popis: {item.Popis,-35}, Příkon: {item.Prikon,-15}");
                }

            }
        }

        /// <summary>Převod stringu na enum</summary>
        private static void StringToEnum(IGrouping<string, Zarizeni> skupina) {
            foreach(var ukol in skupina) {
                if(Enum.TryParse<Druhy>(ukol.Druh, true, out var priorita)) {
                    ukol.DruhEnum = priorita;
                }
                else {
                    ukol.DruhEnum = Druhy.Nic; // nebo jiná výchozí hodnota
                }
            }
        }
    }
}
