using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Seznam
{
    public class ElektroLoad
    {

        /// <summary>Koirování parametrů</summary>
        public void Elektro()
        {
            string cesta = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
            var PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = new ExcelLoad().LoadDataExcel(cesta, PouzitProTabulku, "Motory500V", 2, []);

            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx";
            //var TextPole = new string[] { "Tag", "Příkon", "Měnič", "Balená Jednotka", "Popis", "PID"
            var TextPole = new string[] { "Tag", "Popis", "Prikon", "Menic", "BalenaJednotka", "PID" };
            PouzitProTabulku = new int[] { 3, 18, 21, 1, 7, 2 };
            var Nova = new ExcelLoad().LoadDataExcel(cesta, PouzitProTabulku, "M_equipment_list", 7, TextPole);

            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_ a_spotrebicu_rev6_ELE.xlsx";
            TextPole = new string[] { "Tag", "HP", "Měnič", "Proud", "Delka", "AWG", "BalenaJednotka", "Popis", "Rozvadec", "RozvadecCislo", "PruzezMM2" };
            PouzitProTabulku = new int[] { 5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45 };
            var Stara = new ExcelLoad().LoadDataExcel(cesta, PouzitProTabulku, "M_equipment_list", 7, TextPole);

            // Najít chybějící klíče v obou seznamech
            // Najdeme položky v Nova, které nejsou v Stara.
            var missingFromList2 = FindMissingKeys(Nova, Stara);

            // Najdeme položky v Stara, které nejsou v Nova.
            var missingFromList1 = FindMissingKeys(Stara, Nova);

            //Cesta Excel pro změny 
            cesta = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03\BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx";
            //Hledání shody radku excelu s polem SadaUpraveno --- PouzitProTabulku -> první je kryterium
            PouzitProTabulku = new int[] { 3, 18 };

            //zapis do buněk.
            //var PouzitProZapis = new int[] { 56, 57 };
            var PouzitProZapis = new int[] { 59, 65, 66, 61, 63, 64 };
            new ExcelApp().ExcelSaveSloupec(cesta, PouzitProZapis, zalozka: "M_equipment_list", PouzitProTabulku, Stara);
            Console.Write("\nFunguje --- ExelSaveSlopec ");

            Console.WriteLine(missingFromList2);
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
            //volba kdy jsem doma a kdy v práci - volba dle nazvu PC
            bool Doma = true;
            string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            if (Environment.MachineName == "MARTIN")
            {
                basePath = @"D:\Tractebel\2024.09.03";
                Doma = false;
            }

            var ExcelApp = new ExcelApp();
            var Load = new ExcelLoad();

            //Načtení json z Milanového seznamu čerpadel
            var Pumps = new List<Pump>();
            if (Doma)
            {
                string cestaPump = @"U:\Elektro\mcsato\Zakázky\Natron\pumps.json";
                if (File.Exists(cestaPump))
                { Pumps = Pump.Load(cestaPump); }

            }

            //vytvoření nebo otevření dokumentu elekro
            var cesta = Path.Combine(basePath, "Seznam.xlsx");
            Exc.Worksheet xls = ExcelApp.ExcelElektro(cesta);
            Exc.Workbook doc = xls.Parent;

            Console.WriteLine("\n----------------------------");
            Console.WriteLine("Načtení seznamu ze zdroje A/N");
            if (Console.ReadKey().Key == ConsoleKey.A)
            {
                //načtení základní infomací pro seznam Elektro dle čísel jednotlivých sloupců
                string cesta1 = Path.Combine(basePath, @"BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx");
                //TextPole = new string[] { "Tag", "HP", "Měnič", Proud, Delka,    AWG  "Balená Jednotka", "Popis",  Rozvaděč,   RozvaděčCislo , mm2 };
                //var PouzitProTabulku = new int[] { 5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45 };

                var TextPole = new string[] { "Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "Proud500", "HP", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
                var PouzitProTabulku1 = new int[] { 3, 2, 7, 18, 1, 21, 59, 56, 60, 63, 64, 61, 62, 65, 66 };
                //převod                           3, 2, 7, 18, 1, 21, A, HP,  A, mm2, AWG, m,  ft  mcc cislo
                //var Kotrola                    { 1,  2,     3,       4,           5,              6,      7,          8,      9,        10,   11,     12,         13,         14,     15 };
                var Stara = Load.LoadDataExcel(cesta1, PouzitProTabulku1, "M_equipment_list", 7, TextPole);
                //var Zakalad = Load.LoadDataExcelTrida(cesta, PouzitProTabulku, "M_equipment_list", 7, TextPole);

                //Vytvoření nadpisů
                var Souradnice = ExcelApp.Nadpis(xls);

                //Formátování nadpisů
                ExcelApp.NadpisSet(xls, Souradnice);

                //uložení základní seznam zařízení dle seznamu Stara
                //var TabulkuProPeevod = new int[] { 1, 2, 3, 4,  5, 6,  7, 8,   9,  10,  11, 12, 13,  14,  15 };
                new ExcelApp().ExcelSaveList(xls, Stara);

                if (Doma)
                {
                    //naštení tabulky proudů 
                    cesta = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
                    PouzitProTabulku1 = new int[] { 1, 2, 3 };
                    var Motory500 = Load.LoadDataExcel(cesta, PouzitProTabulku1, "Motory500V", 2, []);
                    //doplnění tabulky proudů rabulky Excel
                    ExcelApp.ExcelSaveProud(xls, Motory500);
                }

                //doplnění vzorců doExel
                ExcelApp.ExcelSaveVzorce(xls, Stara.Count);

                cesta = Path.Combine(basePath, @"BLUECHEM_seznam_stroju_ a_spotrebicu_rev6_ELE.xlsx");
                //TextPole = new string[] { "Tag", "HP", "Měnič", Proud, Delka,    AWG  "Balená Jednotka", "Popis",  Rozvaděč,   RozvaděčCislo , mm2 };
                PouzitProTabulku1 = new int[] { 5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45 };
                var Delka = Load.LoadDataExcel(cesta, PouzitProTabulku1, "M_equipment_list", 7, []);
                //doplnění kabelů z //delka  //awg  //mm2
                //---- v budoucnu kontrola pokud by něco chybělo
                //ExcelApp.ExcelSaveKabel(xls, Delka);

                //doplnění rozvaděčů mcc cislo
                //---- v budoucnu kontrola pokud by něco chybělo
                //new ExcelApp().ExcelSaveRozvadec(xls, Delka);

                //Testovací kod
                //new ExcelApp().PridatTextyTestovani(xls);
            }
            Console.WriteLine("Probíhá načítaní kabelů");
            //Načti seznam zařízení z vytvořeného seznamu zařízení elektro 
            //TextPole = new string[] { "Tag", "PId" "Jmeno", "kW", "BalenaJednotka", "Menic" "Proud500",  "HP"  "Proud480", "mm2" , "AWG" , "Delkam",  Delkaft,     MCC ,  cisloMCC  };
            var PouzitProTabulku = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };

            //v poli jsou čísla posunuty o jedničku
            var PoleData = ExcelApp.ExcelLoadWorksheet(xls, PouzitProTabulku);

            //Úprava načteného listu seznamu zařízení elektro 
            PoleData = new KabelList().Kabely(PoleData);

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

            Console.Write($"\nPocet zaznamu:{unikatniZaznamy.Count()}");

            var Soucet = new List<List<string>>();
            foreach (var item in unikatniZaznamy)
            {
                // Filtruj záznamy podle kritérií a proveď součet
                var soucet = PoleData
                    .Where(z => z[4] == item[4] && z[5] == item[5] && z[6] == item[6]) // Filtrace podle kritérií
                    .Sum(sum => double.TryParse(sum[18], out double hodnota) ? hodnota : 0); // Převod textu na číslo a součet

                Console.Write($"\nzaznamu: {item[4]},{item[5]},{item[6]}, Soucet = {soucet}");
                string[] xx = [item[4], item[5], item[6], soucet.ToString(), (soucet * 3.29).ToString()];
                Soucet.Add([.. xx]);
            }

            //celkový kontrolní součet
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
            new ExcelApp().ExcelQuit(doc);
        }
    }
}
