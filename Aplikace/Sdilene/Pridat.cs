using Aplikace.Excel;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Sdilene
{
    public static class Pridat
    {
        /// <summary>výpočet proudu </summary>
        public static List<Zarizeni> AddProud(this List<Zarizeni> pole)
        {
            // Přidání vlastnosti "Proud" do každého zařízení
            var nove = new List<Zarizeni>();
            foreach (var item in pole)
            {
                if (double.TryParse(item.Napeti, out double U))
                    if (double.TryParse(item.Prikon, out double kW))
                        item.Proud = (kW * 1000 / (Math.Sqrt(3) * U * 0.85) ).ToString("F2");
                nove.Add(item);
            }
            return nove;
        }

        /// <summary>Pridání délky kabelu </summary>
        public static List<Zarizeni> AddKabel(this List<Zarizeni> pole, int delka = 100)
        {
            // Přidání vlastnosti "Proud" do každého zařízení
            var nove = new List<Zarizeni>();
            foreach (var item in pole)
            {
                item.Delka = delka.ToString();
                nove.Add(item);
            }
            return nove;
        }

        public static void Soucet(Workbook doc, List<List<string>> PoleData)
        {
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

            //Celkový kontrolní součet
            Soucet.Add([]);
            var celek = PoleData.Sum(x => double.TryParse(x[18], out double hodnota) ? hodnota : 0); // Převod textu na číslo a součet
            string[] xx1 = ["", "", "", celek.ToString()];
            Soucet.Add([.. xx1]);

            //nová záložka
            var ExcelApp = new ExcelApp();
            ExcelApp.PridatNovyList("Seznam");
            ExcelApp.Nadpis("A1:C1", "Označeni", Soucet);
            ExcelApp.Nadpis("D1:D1", "Délka", Soucet);
            ExcelApp.Xls.Range["D2"].Value = "[m]";
            ExcelApp.ExcelSaveTable(Soucet, 3);

            //xls.Cells[Soucet.Count + 1, 4].Formula = xxx může nastat chyba.
            ExcelApp.Xls.Cells[Soucet.Count + 1, 4].FormulaLocal = $"=SUMA(D3:D{Soucet.Count})"; // SUMAE{i}*500/480";
        }
    }
}
