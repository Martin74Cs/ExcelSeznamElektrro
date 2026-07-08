﻿using Knihovna;
using Knihovna.Excel;
using Knihovna.Sdilene;
using Knihovna.Tridy;

namespace Aplikace.Seznam
{
    public class ElektroLoad
    {

        /// <summary>Koirování parametrů</summary>
        public static void Elektro()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr souboru s motory (např. Motory500V.xlsx)...");
            Console.ResetColor();
            string? cestaMotory = Soubory.ShowOpenFileDialog("Excel soubory (*.xls;*.xlsx)|*.xls;*.xlsx");
            if (string.IsNullOrEmpty(cestaMotory))
            {
                Console.WriteLine("Výběr souboru s motory byl stornován. Proces ukončen.");
                return;
            }

            var PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = ExcelLoad.LoadDataExcel(cestaMotory, PouzitProTabulku, "Motory500V", 2);
            Motory500.Vypis();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr NOVÉHO seznamu strojů (např. rev7)...");
            Console.ResetColor();
            string? cestaNova = Soubory.ShowOpenFileDialog("Excel soubory (*.xls;*.xlsx)|*.xls;*.xlsx");
            if (string.IsNullOrEmpty(cestaNova))
            {
                Console.WriteLine("Výběr nového seznamu byl stornován. Proces ukončen.");
                return;
            }

            PouzitProTabulku = [3, 18, 21, 1, 7, 2];
            var Nova = ExcelLoad.LoadDataExcel(cestaNova, PouzitProTabulku, "M_equipment_list", 7);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr STARÉHO seznamu strojů (např. rev6)...");
            Console.ResetColor();
            string? cestaStara = Soubory.ShowOpenFileDialog("Excel soubory (*.xls;*.xlsx)|*.xls;*.xlsx");
            if (string.IsNullOrEmpty(cestaStara))
            {
                Console.WriteLine("Výběr starého seznamu byl stornován. Proces ukončen.");
                return;
            }

            PouzitProTabulku = [5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45];
            var Stara = ExcelLoad.LoadDataExcel(cestaStara, PouzitProTabulku, "M_equipment_list", 7);

            // Najít chybějící klíče v obou seznamech
            // Najdeme položky v Nova, které nejsou v Stara.
            var MissingFromList2 = FindMissingKeys(Nova, Stara);
            Console.WriteLine(MissingFromList2);

            // Najdeme položky v Stara, které nejsou v Nova.
            var MissingFromList1 = FindMissingKeys(Stara, Nova);
            Console.WriteLine(MissingFromList1);

            // Zápis změn zpět do nového souboru
            //Hledání shody radku excelu s polem SadaUpraveno --- PouzitProTabulku -> první je kryterium
            PouzitProTabulku = [3, 18];

            //zapis do buněk.
            var PouzitProZapis = new int[] { 59, 65, 66, 61, 63, 64 };
            new ExcelApp().ExcelSaveSloupec(cestaNova, PouzitProZapis, zalozka: "M_equipment_list", PouzitProTabulku, Stara);
            Console.Write("\nFunguje --- ExelSaveSlopec ");

            Console.WriteLine(MissingFromList2);
        }

        static List<List<string>> FindMissingKeys(List<List<string>> sourceList, List<List<string>> compareList)
        {
            // Vytvoříme množinu klíčů z compareList
            var compareKeys = new HashSet<string>(compareList.Select(x => x[0]));

            // Najdeme položky v sourceList, které nejsou v compareKeys
            return [.. sourceList.Where(x => !compareKeys.Contains(x[0]))];
        }

        public static void NovyExcel()
        {
            //Volba kdy jsem doma a kdy v práci - volba dle nazvu PC
            bool Doma = true;
            string basePath = @"c:\a\Natron\2024.09.03";
        
            if (Environment.MachineName == "MARTIN")
            {
                basePath = @"D:\Tractebel\2024.09.03";
                Doma = false;
            }
            var cestaXls = Path.Combine(basePath, "Seznam.xlsx");

            var ExcelApp = new ExcelApp();
            var Load = new ExcelLoad();


            //Načtení json z Milanového seznamu čerpadel
            var Pumps = new List<Pump>();
            if (Doma)
            {
                string cestaPump = @"U:\Elektro\mcsato\Zakázky\Natron\pumps.json";
                if (File.Exists(cestaPump))
                    Pumps = Pump.Load(cestaPump); 
            }

            Console.WriteLine("\n----------------------------");
            Console.WriteLine("Načtení seznamu ze zdroje A/N");
            if (Console.ReadKey().Key == ConsoleKey.A)
            {
                //načtení základní infomací pro seznam Elektro dle čísel jednotlivých sloupců
                string cesta1 = Path.Combine(basePath, @"BLUECHEM_seznam_stroju_a_spotrebicu_rev7_ELE_MC.xlsx");

                var TextPole = new string[] { "Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "Proud500", "HP", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
                var PouzitProTabulku1 = new int[] { 3, 2, 7, 18, 1, 21, 59, 56, 60, 63, 64, 61, 62, 65, 66 };
                //převod                           3, 2, 7, 18, 1, 21, A, HP,  A, mm2, AWG, m,  ft  mcc cislo
                var Stara = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku1, "M_equipment_list", 7);

                //vytvoření nebo otevření dokumentu elekro
                var cesta = Path.Combine(basePath, "Seznam.xlsx");
                ExcelApp = new ExcelApp(cesta);

                //Vytvoření nadpisů
                var Souradnice = ExcelApp.Nadpisy([.. Nadpis.DataEn()]);

                //Formátování nadpisů
                ExcelApp.NadpisSet(Souradnice);

                //uložení základní seznam zařízení dle seznamu Stara
                ExcelApp.ExcelSaveList(Stara);

                if (Doma)
                {
                    //naštení tabulky proudů 
                    cesta = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
                    PouzitProTabulku1 = [1, 2, 3];
                    var Motory500 = ExcelLoad.LoadDataExcel(cesta, PouzitProTabulku1, "Motory500V", 2);
                    //doplnění tabulky proudů rabulky Excel
                    ExcelApp.ExcelSaveProud(Motory500);
                }

                //doplnění vzorců doExel
                ExcelApp.ExcelSaveVzorce(Stara.Count);

                cesta = Path.Combine(basePath, @"BLUECHEM_seznam_stroju_ a_spotrebicu_rev6_ELE.xlsx");
                PouzitProTabulku1 = [5, 38, 23, 41, 43, 46, 3, 9, 47, 48, 45];

                //doplnění kabelů z //delka  //awg  //mm2
                //---- v budoucnu kontrola pokud by něco chybělo

                //doplnění rozvaděčů mcc cislo
                //---- v budoucnu kontrola pokud by něco chybělo

                //Testovací kod
            }
            else
            { 
                //vytvoření nebo otevření dokumentu elekro
                ExcelApp = new ExcelApp(cestaXls);
            }

            Console.WriteLine("Probíhá načítaní kabelů");
            //Načti seznam zařízení z vytvořeného seznamu zařízení elektro 
            var PouzitProTabulku = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };

            //V poli jsou čísla posunuty o jedničku
            var PoleData = ExcelApp.ExcelLoadWorksheet(PouzitProTabulku);

            //Úprava načteného listu seznamu zařízení elektro 

            //Nová záložka
            ExcelApp.GetSheet("Kabely");

            //doplnení nadpisu
            ExcelApp.ExcelSaveNadpis(PoleData);

            //do Excel vyplní od radku 3 data data z PoleData mělo by se jednat o seznam kabelů
            ExcelApp.KabelyToExcel(PoleData, 3);

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
                Soucet.Add([.. xx]);
            }

            //celkový kontrolní součet
            Soucet.Add([]);
            var celek = PoleData.Sum(x => double.TryParse(x[18], out double hodnota) ? hodnota : 0); // Převod textu na číslo a součet
            string[] xx1 = ["", "", "", celek.ToString()];
            Soucet.Add([.. xx1]);

            //nová záložka
            ExcelApp.GetSheet("Seznam");
            ExcelApp.Nadpis("A1:C1", "Označeni", Soucet.Count);
            ExcelApp.Nadpis("D1:E1", "Délka", Soucet.Count);
            ExcelApp.Xls.Range["D2"].Value = "[m]";
            ExcelApp.Xls.Range["E2"].Value = "[ft]";
            ExcelApp.KabelyToExcel(Soucet, 3);

            ExcelApp.Xls.Cells[Soucet.Count + 1, 4].Formula = $"=SUMA(D3:D{Soucet.Count})"; // SUMAE{i}*500/480";
            ExcelApp.ExcelQuit(cestaXls);
        }
    }
}

