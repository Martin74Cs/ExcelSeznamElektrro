using Aplikace.Excel;
using Aplikace.Seznam;
using Aplikace.Tridy;
using System.Net.Sockets;
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
            //var Stara = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku1, "M_equipment_list", 7, TextPole);
            var Stara = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku1, "M_equipment_list", 7, TextPole);
            //var Zakalad = Load.LoadDataExcelTrida(cesta, PouzitProTabulku, "M_equipment_list", 7, TextPole);
            
            //vytvoření nebo otevření dokumentu elekro
            var cesta = Path.Combine(basePath, "Seznam.xlsx");
            xls = ExcelApp.ExcelElektro(cesta);
            doc = xls.Parent;

            //Vytvoření nadpisů
            var Souradnice = ExcelApp.Nadpis(xls);

            //Formátování nadpisů
            ExcelApp.NadpisSet(xls, Souradnice);

            //uložení základní seznam zařízení dle seznamu Stara
            //var TabulkuProPeevod = new int[] { 1, 2, 3, 4,  5, 6,  7, 8,   9,  10,  11, 12, 13,  14,  15 };
            ExcelApp.ExcelSaveList(xls, Stara);

            //doplnění vzorců doExel
            ExcelApp.ExcelSaveVzorce(xls, Stara.Count);
        
            //else
            //{ 
            //    //vytvoření nebo otevření dokumentu elekro
            //    cesta = Path.Combine(basePath, "Seznam.xlsx");
            //    xls = ExcelApp.ExcelElektro(cesta);
            //    doc = xls.Parent;
            //}

            Console.WriteLine("Probíhá načítaní kabelů");
                    //Načti seznam zařízení z vytvořeného seznamu zařízení elektro 
                    //TextPole = new string[] { "Tag", "PId" "Jmeno", "kW", "BalenaJednotka", "Menic" "Proud500",  "HP"  "Proud480", "mm2" , "AWG" , "Delkam",  Delkaft,     MCC ,  cisloMCC  };
                    var PouzitProTabulku = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };

            //v poli jsou čísla posunuty o jedničku
            var PoleData = ExcelApp.ExcelLoadWorksheet(xls, PouzitProTabulku);

            //Úprava načteného listu seznamu zařízení elektro 
            PoleData = KabelList.Kabely(PoleData);

            //Nová záložka
            xls = ExcelApp.PridatNovyList(doc, "Kabely");

            //doplnení nadpisu
            ExcelApp.ExcelSaveNadpis(xls, PoleData);

            //do Excel vyplní od radku 3 data data z PoleData mělo by se jednat o seznam kabelů
            ExcelApp.ExcelSaveTable(xls, PoleData, 3);

            //vyzváření seznamu kabelů podle krytérii
            // Použití GroupBy k získání unikátních záznamů na základě tří kritérií
            var unikatniZaznamy = PoleData
                //4. kabel CYKY, 5. počet vodiců, 6. Přůřez
                .GroupBy(z => new { Krit1 = z[4], Krit2 = z[5], Krit3 = z[6] }) // Skupinování podle kritérií
                .Select(g => g.First()) // Vybereme první záznam z každé skupiny
                .ToList();

            Console.Write($"\nPocet zaznamu:{unikatniZaznamy.Count}");

            var Soucet = new List<List<string>>();
            foreach (var item in unikatniZaznamy)
            {
                // Filtruj záznamy podle kritérií a proveď součet
                var soucet = PoleData
                    .Where(z => z[4] == item[4] && z[5] == item[5] && z[6] == item[6]) // Filtrace podle kritérií
                    .Sum(sum => double.TryParse(sum[18], out double hodnota) ? hodnota : 0); // Převod textu na číslo a součet

                Console.Write($"\nzaznamu: {item[4]},{item[5]},{item[6]}, Soucet = {soucet}");
                            string[] xx = [item[4], item[5], item[6], soucet.ToString(), (soucet * 3.29).ToString()];
                Soucet.Add([..xx]);
            }

            //Celkový kontrolní součet
            Soucet.Add([]);
            var celek = PoleData.Sum(x => double.TryParse(x[18], out double hodnota) ? hodnota : 0); // Převod textu na číslo a součet
            string[] xx1 = ["", "", "", celek.ToString()];
            Soucet.Add([.. xx1]);

            //nová záložka
            xls = ExcelApp.PridatNovyList(doc, "Seznam");
            ExcelApp.Nadpis(xls, "A1:C1", "Označeni", Soucet);
            ExcelApp.Nadpis(xls, "D1:E1", "Délka", Soucet);
            xls.Range["D2"].Value = "[m]";
            xls.Range["E2"].Value = "[ft]";
            ExcelApp.ExcelSaveTable(xls, Soucet, 3);


            xls.Cells[Soucet.Count + 1, 4].Formula = $"=SUMA(D3:D{Soucet.Count})"; // SUMAE{i}*500/480";
            ExcelApp.ExcelQuit(doc);

        }
    }
}
