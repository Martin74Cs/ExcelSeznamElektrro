using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using System.Reflection;
using System.Runtime.InteropServices;
using static Aplikace.Tridy.Motor;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Upravy
{
    public class LigthChem
    {
        public static string BasePath {
            get
            {
                if (Environment.UserDomainName == "D10")
                    return @"c:\a\LightChem\Elektro\";
                else
                    return @"G:\Můj disk\Elektro";
            }
        }
        public static string GooglePath
        {
            get
            {
                if (Environment.UserDomainName == "D10")
                    return @"E:\Můj disk\Elektro";
                else
                    return @"G:\Můj disk\Elektro";
            }
        }
        public static void Hlavni()
        {
            //var xxx = new Zarizeni();
            //xxx.Vypis();

            string basePath;
            if (Environment.UserDomainName == "D10")
            {
                Console.WriteLine("Jsem v práci");
                basePath = @"c:\a\LightChem\Elektro\";
            }
            else { 
                Console.WriteLine("Jsem doma na Terase");
                basePath = @"G:\Můj disk\Elektro";
                }

            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            string filename = "Seznam.xlsx";
            if (!Directory.Exists(basePath))
                Directory.CreateDirectory(basePath);
            //string cesta1 = Path.Combine(basePath, @"BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx");

            ////načtení základní infomací pro seznam Elektro dle čísel jednotlivých sloupců
            //string[] TextPole =     ["Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "Proud500", "HP", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC"];
            //int[] PouzitProTabulku1 = [3,   2,      7,      18,         1,              21,         59,     56,     60,         63,     64,     61,     62,         65,     66];
            //var Stara = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku1, "M_equipment_list", 7, TextPole);

            //string cesta1 = Path.Combine(basePath, @"N78020_Consumer_List.xls");
            //var Stara = ExcelLoad.DataExcel(cesta1, "Seznam", 4);
            string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            var Stara = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            //Výpočet položky proud
            Stara.AddProud();
            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Prevod.SaveToCsv(Stara, Path.ChangeExtension(cesta1, ".csv"));

            //vytvoření nebo otevření dokumentu elektro
             var cesta = Path.Combine(basePath, filename);
            var ExcelApp = new ExcelApp(cesta);
            //var (App, Doc ,Xls) = ExcelApp.ExcelElektro(cesta);
            //ExcelApp.ExcelElektro(cesta);
            
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
            ExcelApp.ClassToExcel(Row: 3, Stara);
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
            var PoleData = KabelList.Kabely(Stara);

            //Nová záložka
            ExcelApp.PridatNovyList("Kabely");

            //Doplnení nadpisu a ramecku
            ExcelApp.ExcelSaveNadpis(PoleData);

            //do Excel vyplní od radku 3 data data z PoleData mělo by se jednat o seznam kabelů
            //Dlouho dočasně vypnuto
            //ExcelApp.ExcelSaveTable(PoleData, 3);

            //vyzváření seznamu kabelů podle krytérii
            Pridat.Soucet(ExcelApp, PoleData);
            
            //var Proces =  Soubory.GetExcelProcess(ExcelApp.App);
            if (!File.Exists(cesta))
                ExcelApp.Doc.SaveAs(cesta);
            else
            { 
                ExcelApp.Doc.Save();
                //if(!Soubory.IsFileLocked(cesta))
                    // Zavření bez uložení
                    //ExcelApp.Doc.Close(false);
            }
            //ExcelApp.App.Quit();

            //if (File.Exists(cesta))
            //    File.Delete(cesta);
            //doc.SaveAs(cesta);
            ExcelApp.ExcelQuit();
        }

        public static void KabelyAdd()
        {
            string basePath = @"G:\Můj disk\Elektro";
            //string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.json");
            var Target = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            var KabelCu = Soubory.LoadJsonList<KabelVse>(Cesty.CuJson)
                .Where(x => x.Name.Contains("CYKY", StringComparison.OrdinalIgnoreCase) && x.Deleni == "4")
                .OrderBy(x => x.IzAE).ToList();
            KabelCu.SaveJsonList(Path.ChangeExtension(cesta1, ".txt"));

            var KabelAL = Soubory.LoadJsonList<KabelVse>(Cesty.AlJson).OrderByDescending(x => x.IzAE );

            //item nevím co to je
            var properties = typeof(Zarizeni).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => p.CanWrite && p.Name != "Item")
                                 .ToList();

            foreach (var target in Target)
            {
                foreach (var prop in properties)
                {
                    //var value = prop.GetValue(target);
                    var JedenKabel = KabelCu.FirstOrDefault(x => x.MaxProud > Convert.ToDouble(target.Proud));
                    if (JedenKabel == null) continue;

                    //prop.SetValue(target, value);
                    //target.GetType().GetProperty("PruzezMM2").SetValue(target, JedenKabel.SLmm2.Tostring());
                    target.PruzezMM2 = JedenKabel.SLmm2.ToString();
                    target.Deleni = JedenKabel.Deleni;
                    target.Kabel = JedenKabel;
                }

            }
            Console.WriteLine("Save Json");
            Target.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
        }

        public static void PrevodCsvToJson()
        {
            string basePath = @"G:\Můj disk\Elektro";
            //string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.json");
            var Target = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            string filename = @"N92120_Seznam_stroju_zarizeni_250311_250407.csv";
            string cesta = Path.Combine(basePath, filename);
            if (!File.Exists(cesta)) return;
            var Source = Soubory.LoadFromCsv<Zarizeni>(cesta);
            
            Prevod.UpdateCsvToJson(Source, Target );

            Target.SaveJsonList(cesta1);
            Prevod.SaveToCsv(Target, Path.ChangeExtension(cesta1, ".csv"));
        }

        public static void VyvoritFMKM()
        {
            var cesta = Environment.ProcessPath;            // Získá úplnou cestu ke spuštěnému procesu
            var dir = Path.GetDirectoryName(cesta);         // Získá adresář, kde je spustitelný soubor

            string basePath = Path.Combine(GooglePath, "Data");
            //string CestaFM = @"C:\VSCode\ExcelSeznamElektrro\Aplikace\Data\FM.csv";
            //string CestaKM = @"C:\VSCode\ExcelSeznamElektrro\Aplikace\Data\KM.csv";
            string CestaFM = Path.Combine(basePath, "FM.csv");
            string CestaKM = Path.Combine(basePath, "KM.csv"); 

            var FM = Soubory.LoadFromCsv<Menic>(CestaFM);
            Console.WriteLine($"Pocet menicu: {FM.Count}");
            
            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            FM.SaveJsonList(Path.ChangeExtension(CestaFM, ".json"));
            Console.WriteLine($"Měniče uloženy jako Json");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        public static void VyvoritMotor()
        {
            string basePath = Path.Combine(GooglePath, "Data" , "Motory");
            string CestaMotor = Path.Combine(basePath, "Motory.csv");
            var Motor = Soubory.LoadFromCsv<Motor>(CestaMotor);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            string CestaMotor3000 = Path.Combine(basePath, "Motory3000.csv");
            var Motor3000 = Soubory.LoadFromCsv<Motor>(CestaMotor);
            Console.WriteLine($"Pocet motorů: {Motor3000.Count}");

            Motor.AddRange(Motor3000);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            Motor.SaveJsonList(Path.ChangeExtension(CestaMotor, ".json"));
            Console.WriteLine($"Motory uloženy jako Json");
        }
    }
}
