using Aplikace.Excel;
using Aplikace.Sdilene;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Seznam
{
    public class ElektroLoad
    {
        public void Elektro()
        { 
            string cesta = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
            var PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = new ExcelLoad().LoadDataExcel(cesta, PouzitProTabulku, "Motory500V", 2);

            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx";
            //var TextPole = new string[] { "Tag", "Příkon", "Měnič", "Balená Jednotka", "Popis", "PID" };
            PouzitProTabulku = new int[] { 3, 18, 21, 1, 7, 2 };
                var Nova = new ExcelLoad().LoadDataExcel(cesta, PouzitProTabulku, "M_equipment_list", 7);

            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_ a_spotrebicu_rev6_ELE.xlsx";
            //TextPole = new string[] { "Tag", "HP", "Měnič", Proud, Delka,    AWG  "Balená Jednotka", "Popis",  Rozvaděč,   RozvaděčCislo , mm2 };
            PouzitProTabulku = new int[] { 5,   38,     23,     41,     43,     46,      3,             9,          47,         48 ,        45   };
            var Stara = new ExcelLoad().LoadDataExcel(cesta, PouzitProTabulku, "M_equipment_list", 7);

            // Najít chybějící klíče v obou seznamech
            // Najdeme položky v Nova, které nejsou v Stara.
            var missingFromList2 = FindMissingKeys(Nova, Stara);

            // Najdeme položky v Stara, které nejsou v Nova.
            var missingFromList1 = FindMissingKeys(Stara, Nova);

            //Cesta Excel pro změny změny
            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx";
            //Hledání shody radku excelu s polem SadaUpraveno --- PouzitProTabulku -> první je kryterium
            PouzitProTabulku = new int[] { 3, 18 };

            //zapis do buněk.
            //var PouzitProZapis = new int[] { 56, 57 };
            var PouzitProZapis = new int[] { 59, 65, 66, 61, 63, 64 };
            new ExcelApp().ExcelSaveSloupec(cesta, PouzitProZapis, zalozka: "M_equipment_list", PouzitProTabulku, Stara);
            Console.Write("\nFunguje --- ExelSaveSlopec ");

            Console.WriteLine(missingFromList2);

            //var Ex = new Aplikace.Excel.ExcelApp();
            //Exc.Workbook dok = Ex.DokumetExcel(cesta);
            //if (dok == null) return;
            //Console.Write("\nsheet" + dok.Name + " , pocet=" + dok.Worksheets.Count);

            //Exc.Worksheet sheet = Ex.GetSheet(dok, "M_equipment_list");
            //if (sheet == null) { Console.WriteLine("\nPožadovaný list nebyl nalezen!"); return; }
            //Console.Write("\nByl nalezen list - " + sheet.Name);

            ////int[] Sloupec = new int[] { 4, 8, 19, 22 };
            //int[] Sloupec = [3, 8, 19, 22];
            //int Nadpis = 5;
            //var TableDate = Ex.GetTable(sheet, rowNadpis: Nadpis, Sloupec);
            //Console.Write(" ... OK");
            ////konec excel
            //Ex.ExcelQuit(dok);

            //kontrola velikosti
            //Console.Write("\nTabulka má {0} radků.", TableDate.Rows.Count);
            //if (TableDate.Rows.Count < 1) return;
        }

        static List<List<string>> FindMissingKeys(List<List<string>> sourceList, List<List<string>> compareList)
        {
            // Vytvoříme množinu klíčů z compareList
            HashSet<string> compareKeys = new HashSet<string>(compareList.Select(x => x[0]));

            // Najdeme položky v sourceList, které nejsou v compareKeys
            return sourceList.Where(x => !compareKeys.Contains(x[0])).ToList();
        }

        public void NovyExcel()
        {
            var ExcelApp = new ExcelApp();
            var Load = new ExcelLoad();

            //načtení základní infomací pro seznam Elektro dle čísel jednotlivých sloupců
            string cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx";
            //TextPole = new string[] { "Tag", "HP", "Měnič", Proud, Delka,    AWG  "Balená Jednotka", "Popis",  Rozvaděč,   RozvaděčCislo , mm2 };
            //var PouzitProTabulku = new int[] { 5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45 };
            //var TextPole = new string[] { "Tag","Popis", "Příkon", "Měnič", "Balená Jednotka", , "PID" };
            var PouzitProTabulku = new int[] { 3,   7,      18,         21,         1                    };
            var Stara = Load.LoadDataExcel(cesta, PouzitProTabulku, "M_equipment_list", 7);

            //vytvoření nebo otevření dokumentu elekro
            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\Seznam.xlsx";
            Exc.Worksheet xls = ExcelApp.ExcelElektro(cesta);
            Exc.Workbook doc = xls.Parent;

            //Vytvoření dapisů
            ExcelApp.Nadpis(xls);
            //Formátování nadpisů
            ExcelApp.NadpisSet(xls);
            //uložení základní seznam zařízení dle seznamu Stara
            new ExcelApp().ExcelSaveList(xls, Stara);

            //naštení tabulky proudů 
            cesta = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
            PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = Load.LoadDataExcel(cesta, PouzitProTabulku, "Motory500V", 2);
            //doplnění tabulky proudů rabulky Excel
            ExcelApp.ExcelSaveProud(xls, Motory500);

            //doplnění vzorců doExel
            ExcelApp.ExcelSaveVzorce(xls);

            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_ a_spotrebicu_rev6_ELE.xlsx";
            //TextPole = new string[] { "Tag", "HP", "Měnič", Proud, Delka,    AWG  "Balená Jednotka", "Popis",  Rozvaděč,   RozvaděčCislo , mm2 };
            PouzitProTabulku = new int[] { 5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45 };
            var Delka = Load.LoadDataExcel(cesta, PouzitProTabulku, "M_equipment_list", 7);
            //doplnění kabelů z //delka  //awg  //mm2
            ExcelApp.ExcelSaveKabel(xls, Delka);

            //doplnění rozvaděčů mcc cislo
            new ExcelApp().ExcelSaveRozvadec(xls, Delka);

            //Testovací kod
            //new ExcelApp().PridatTextyTestovani(xls);

            //Načti seznam zařízení z vytvořeného seznamu zařízení elektro 
            //TextPole = new string[] { "Tag", "Jmeno", "kW", "Zapojeni", "Proud500",  "HP"  "Proud480", "Delkam",  Delkaft, mm2 , AWG ,    MCC ,  cisloMCC  };
            PouzitProTabulku = new int[] { 1,   2,      3,      4,          5,          6,      7,          8,      9,       10,    11,     12,     13 };           
            //v poli jsou čísla posunity o jedničku
            var PoleData = new ExcelApp().ExcelLoadWorksheet(xls, PouzitProTabulku);

            //Uprava našteného listu seznamu zařízení elektro 
            PoleData = new KabelList().Kabely(PoleData);

            //nová záložka
            xls = new ExcelApp().PridatNovyList(doc, "Kabely");

            //doplnení nadpoisu
            new ExcelApp().ExcelSaveNadpis(xls);

            //doplnění rozvaděčů mcc cislo
            new ExcelApp().ExcelSaveTable(xls, PoleData, 3);

            new ExcelApp().ExcelQuit(doc);

        }
    }
}
