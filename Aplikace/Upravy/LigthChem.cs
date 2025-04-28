using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using static Aplikace.Tridy.Motor;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Upravy
{
    public class LigthChem
    {
        public static void Hlavni()
        {
            //var xxx = new Zarizeni();
            //xxx.Vypis();

            //string basePath;
            //if (Environment.UserDomainName == "D10")
            //{
            //    Console.WriteLine("Jsem v práci");
            //    basePath = @"c:\a\LightChem\Elektro\";
            //}
            //else { 
            //    Console.WriteLine("Jsem doma na Terase");
            //    basePath = @"G:\Můj disk\Elektro";
            //}

            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            string filename = "Seznam.xlsx";
            if (!Directory.Exists(Cesty.Elektro))
                Directory.CreateDirectory(Cesty.Elektro);
            //string cesta1 = Path.Combine(basePath, @"BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx");

            ////načtení základní infomací pro seznam Elektro dle čísel jednotlivých sloupců
            //string[] TextPole =     ["Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "Proud500", "HP", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC"];
            //int[] PouzitProTabulku1 = [3,   2,      7,      18,         1,              21,         59,     56,     60,         63,     64,     61,     62,         65,     66];
            //var Stara = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku1, "M_equipment_list", 7, TextPole);

            //string cesta1 = Path.Combine(basePath, @"N78020_Consumer_List.xls");
            //var Stara = ExcelLoad.DataExcel(cesta1, "Seznam", 4);
            string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            var Stara = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            //Výpočet položky proud.
            Stara.AddProud();

            //Přidání typu kabelu.
            AddKabel(Stara);

            //Přidání délky kabelu.
            Stara.AddKabelDelka();

            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Stara.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));

             var cesta = Path.Combine(Cesty.Elektro, filename);
            //vytvoření nebo otevření dokumentu elektro
            var ExcelApp = new ExcelApp(cesta);

            //Nastavení nebo vytvoření záložky
            ExcelApp.GetSheet("Seznam Elektro");
            //Vytvoření nadpisů
            var range = ExcelApp.Nadpisy([.. Nadpis.DataCz()]);

            //Formátování nadpisů
            ExcelApp.NadpisSet(range);

            //if (Stara.Count < 1)
            //{
            //Fake data
            //toto je vzor pro vytvoření tabulky
            //var TextPole = new string[] { "Tag", "PID", "Equipment name", "kW", "BalenaJednotka", "Menic", "Nic", "Power [HP]", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
            //var PouzitProTabulku = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };
            //Stara.Add(["1",     "2",    "3",    "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15"]);
            //}
            //ExcelApp.ExcelSaveClass(xls, Stara);

            //čísla sloupců a nazvy tříd. použíty pouze čísla do 15
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
            //Doplnění vzorců doExel
            //ExcelApp.ExcelSaveVzorce(xls, Stara.Count);

            //else
            //{ 
            //    //vytvoření nebo otevření dokumentu elekro
            //    cesta = Path.Combine(basePath, "Seznam.xlsx");
            //    xls = ExcelApp.ExcelElektro(cesta);
            //    doc = xls.Parent;
            //}

            //Načti seznam zařízení z vytvořeného seznamu zařízení elektro 
            //TextPole = new string[] { "Tag", "PId" "Jmeno", "kW", "BalenaJednotka", "Menic" "Proud500",  "HP"  "Proud480", "mm2" , "AWG" , "Delkam",  Delkaft,     MCC ,  cisloMCC  };
            //var PouzitProTabulku = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };

            //v poli jsou čísla posunuty o jedničku
            //var PoleData = ExcelApp.ExcelLoadWorksheet(xls, PouzitProTabulku);

            //Úprava načteného listu seznamu zařízení elektro 
            Console.WriteLine("Probíhá načítaní kabelů");
            //AddKabel(Stara);

            //Vytvoření poole kabelů pro zápis do Excelu
            var PoleData = KabelList.Kabely(Stara);

            //Nová záložka
            ExcelApp.GetSheet("Kabely");

            //Doplnení nadpisu a ramecku
            ExcelApp.ExcelSaveNadpis(PoleData);

            //do Excel vyplní od radku 3 data data z PoleData mělo by se jednat o seznam kabelů
            //Dlouho dočasně vypnuto
            ExcelApp.KabelyToExcel(PoleData, 3);

            //vyzváření seznamu kabelů podle krytérii
            Pridat.Soucet(ExcelApp, PoleData);
            
            //var Proces =  Soubory.GetExcelProcess(ExcelApp.App);

            //ExcelApp.App.Quit();

            //if (File.Exists(cesta))
            //    File.Delete(cesta);
            //doc.SaveAs(cesta);
            ExcelApp.ExcelQuit(cesta);
        }

        /// <summary>Seznam vývodů pro doplnění </summary>
        public static void AddVyvody()
        {
            //Aktualní seznam vývodů
            string cesta = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.json");
            var Target = ExcelLoad.DataExcel(cesta, "Seznam", 8);

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

            //Načtení vývodů
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

            if (zmena)
            { 
                //Prevod.SaveToCsv(Add, cesta1);
                Add.SaveToCsv(cesta1);
                Add.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            }

            int pocet = Target.Count + 1;
            foreach (var item in Add)
            {
                var Toto = Target.FirstOrDefault(x => x.Apid == item.Apid);
                //existuje tak ho smaž
                if (Toto != null) //continue;
                    Target.Remove(Toto);

                //objekt
                item.Radek = pocet++;
                Target.Add(item);
            }

            Target.SaveJsonList(cesta);
            Target.SaveToCsv(Path.ChangeExtension(cesta, ".csv"));
            //Target.SaveJsonList(Path.ChangeExtension(cesta, ".txt"));
        }

        public static void AddKabely()
        {
            //string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string cesta1 = Path.Combine(Cesty.Lightchem, @"N92120_Seznam_stroju_zarizeni_250311_250407.json");
            var Target = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            AddKabel(Target);

            Target.SaveJsonList(cesta1);
            Target.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));
        }

        public static List<Zarizeni> AddKabel(List<Zarizeni> Target)
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
                    var JedenKabel = KabelCu.FirstOrDefault(x => x.MaxProud > proud * 1.5);
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

        public static void PrevodCsvToJson()
        {
            //string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.json");
            var Target = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            string filename = @"N92120_Seznam_stroju_zarizeni_250311_250407.csv";
            string cesta = Path.Combine(Cesty.Elektro, filename);
            if (!File.Exists(cesta)) return;
            var Source = Soubory.LoadFromCsv<Zarizeni>(cesta);
            
            Prevod.UpdateCsvToJson(Source, Target );

            Target.SaveJsonList(cesta1);
            //Asi zrušit načítám změny s .csv
            //Prevod.SaveToCsv(Target, Path.ChangeExtension(cesta1, ".csv"));
        }
        public static void VyvoritFMKM()
        {
             //var cesta = Environment.ProcessPath;            // Získá úplnou cestu ke spuštěnému procesu
            //var dir = Path.GetDirectoryName(cesta);         // Získá adresář, kde je spustitelný soubor
            VyvoritFM();
            VyvoritKM();
        }

        /// <summary>
        /// Převod seznamu frekvenčních měničů na Json
        /// </summary>
        public static void VyvoritFM()
        {
            string basePath = Path.Combine(Cesty.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv"); 
           
            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary>
        /// Převod seznamu stykačů na Json
        /// </summary>
        public static void VyvoritKM()
        {
            string basePath = Path.Combine(Cesty.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv"); 
           
            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary>
        /// Převod seznamu motorů a motorů3000 na jeden Json
        /// </summary>
        public static void VyvoritMotor()
        {
            string basePath = Path.Combine(Cesty.GooglePath, "Data" , "Motory");
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
