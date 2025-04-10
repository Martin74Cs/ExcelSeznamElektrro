using Aplikace.Excel;
using Aplikace.Tridy;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Upravy
{
    public class LigthChem
    {
        public static void Hlavni()
        {
            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            string basePath = @"c:\a\LightChem\";

            //var ExcelApp = new ExcelApp();
            //var Load = new ExcelLoad();

            Exc.Worksheet xls;
            Exc.Workbook doc;

            //načtení základní infomací pro seznam Elektro dle čísel jednotlivých sloupců
            string cesta1 = Path.Combine(basePath, @"BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx");
            //TextPole = new string[] { "Tag", "HP", "Měnič", Proud, Delka,    AWG  "Balená Jednotka", "Popis",  Rozvaděč,   RozvaděčCislo , mm2 };
            //var PouzitProTabulku = new int[] { 5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45 };

            var TextPole = new string[] { "Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "Proud500", "HP", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
            var PouzitProTabulku1 = new int[] { 3, 2, 7, 18, 1, 21, 59, 56, 60, 63, 64, 61, 62, 65, 66 };
            //převod                            3, 2, 7, 18, 1, 21, A, HP,  A, mm2, AWG, m,  ft  mcc cislo
            //var Kontrola                 { 1,  2,     3,       4,           5,              6,      7,          8,      9,        10,   11,     12,         13,         14,     15 };
            var Stara = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku1, "M_equipment_list", 7, TextPole);
            //var Zakalad = Load.LoadDataExcelTrida(cesta, PouzitProTabulku, "M_equipment_list", 7, TextPole);

            //vytvoření nebo otevření dokumentu elekro
            var cesta = Path.Combine(basePath, "Seznam.xlsx");
            xls = ExcelApp.ExcelElektro(cesta);
            doc = xls.Parent;
            
        }
    }
}
