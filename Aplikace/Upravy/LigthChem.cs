using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Upravy
{
    public class LigthChem
    {
        public static void Hlavni()
        {
            var xxx = new Zarizeni();
            xxx.Vypis();

            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            string basePath = @"c:\a\LightChem\Elektro\";
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
            Prevod.JsonToCsv(Stara, Path.ChangeExtension(cesta1, ".csv"));

            //vytvoření nebo otevření dokumentu elektro
             var cesta = Path.Combine(basePath, filename);
            var ExcelApp = new ExcelApp();
            //var (App, Doc ,Xls) = ExcelApp.ExcelElektro(cesta);
            ExcelApp.ExcelElektro(cesta);

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
            ExcelApp.ExcelSaveTable(PoleData, 3);

            //vyzváření seznamu kabelů podle krytérii
            Pridat.Soucet(PoleData);
            
            //var Proces =  Soubory.GetExcelProcess(ExcelApp.App);
            if (!File.Exists(cesta))
                ExcelApp.Doc.SaveAs(cesta);
            else
            { 
                if(!Soubory.IsFileLocked(cesta))
                    //Doc.Save();
                    // Zavření bez uložení
                    ExcelApp.Doc.Close(false);
            }
            ExcelApp.App.Quit();

            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                if (ExcelApp.Doc != null)
                {
                    Marshal.ReleaseComObject(ExcelApp.Doc);
                    ExcelApp.Doc = null;
                }

                if (ExcelApp.App != null)
                {
                    Marshal.ReleaseComObject(ExcelApp.App);
                    ExcelApp.App = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Soubory.KillExcel(ExcelApp.Process);

            //if (File.Exists(cesta))
            //    File.Delete(cesta);
            //doc.SaveAs(cesta);
            ExcelApp.ExcelQuit();
        }
    }
}
