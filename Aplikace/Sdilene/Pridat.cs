using Aplikace.Excel;
using Aplikace.Tridy;
using Aplikace.Upravy;
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
        public static IEnumerable<Zarizeni> AddProud(this IEnumerable<Zarizeni> pole)
        {
            string Cesta = Path.Combine(Cesty.MotoryJson);
            var Motory = Soubory.LoadJsonList<Motor>(Cesta).Where(x => x.Otacky50 > 2800).OrderBy(x => x.Vykon50).ToList();
            if (Motory.Count < 1) 
            {
                Console.WriteLine($"Nebyly nanačteny motory z {Cesta}");
                return pole;
            }
            // Přidání vlastnosti "Proud" do každého zařízení
            //var nove = new List<Zarizeni>();
            double Cos = 0.95;
            double Pomoc;
            foreach (var item in pole.ToHashSet())
            {
                if (double.TryParse(item.Napeti, out double U) && U != 0 && double.TryParse(item.Prikon, out double kW))
                {

                    if (item.Druh == Zarizeni.Druhy.Rozvadeč.ToString() || item.Druh == Zarizeni.Druhy.Přívod.ToString())
                    {
                        Cos = 1.00;
                    }
                    else { 
                        var JedenMotor = Motory.FirstOrDefault(x => (double)x.Vykon50 >= kW);
                        if (JedenMotor != null)
                        {
                            Cos = JedenMotor.Ucinik50;
                        
                            //přidání motoru do třídy
                            item.Motor = JedenMotor;

                            //přidání příkonu motoru do třídy
                            item.PrikonStroj = item.Prikon;

                            //Uprava příkonu dle velikosti motoru
                            item.Prikon = JedenMotor.Vykon50.ToString("F2");   
                        }
                    }
                    //Pokud je napětí větší než 250V, použijeme vzorec pro třífázový proud
                    if (U > 250)
                        Pomoc = kW * 1000 / (Math.Sqrt(3) * U * Cos);
                    else
                        Pomoc = kW * 1000 / (U * Cos);

                    //zaokrouhluje na dvě desetinná místa (ne ořezává).
                    item.Proud = Pomoc.ToString("F2");
                    var CosString = Pomoc.ToString("F2");
                    Console.WriteLine($"Proud: {item.Proud}, cos: {CosString}");
                }
                else
                { 
                    Console.WriteLine($"Na {item.Radek} Chybí napětí nabo příkon - Apid={item.Apid}");
                }
                //nove.Add(item);
            }
            return pole;
        }

        /// <summary>Pridání délky kabelu </summary>
        public static void AddKabelDelka(this List<Zarizeni> pole, double delka = 100)
        {
            // Přidání vlastnosti "Proud" do každého zařízení
            //var nove = new List<Zarizeni>();
            for (int i = 0; i < pole.Count; i++)
            {
                pole[i].Delka = delka;
                pole[i].Delkaft = delka * 3.28;
            }
            //return pole;
        }

        public static void Soucet(ExcelApp ExcelApp, List<List<string>> PoleData, string SheetName)
        {
            // Použití GroupBy k získání unikátních záznamů na základě tří kritérií
            var unikatniZaznamy = PoleData
                //4. kabel CYKY, 5. počet vodiců, 6. Přůřez
                .GroupBy(z => new { Krit1 = z[4], Krit2 = z[5], Krit3 = z[6] }) // Skupinování podle kritérií
                .Select(g => g.First()) // Vybereme první záznam z každé skupiny
                .ToList();

            Console.Write($"\nPocet zaznamu:{unikatniZaznamy.Count}");

            // Vytvoření seznamu pro součty
            var Soucet = new List<List<string>>();
            foreach (var item in unikatniZaznamy)
            {
                // Filtruj záznamy podle kritérií a proveď součet
                var soucet = PoleData
                    .Where(z => z[4] == item[4] && z[5] == item[5] && z[6] == item[6]) // Filtrace podle kritérií
                    .Sum(sum => double.TryParse(sum[18], out double hodnota) ? hodnota : 0); // Převod textu na číslo a součet

                Console.Write($"\nzaznamu: {item[4]},{item[5]},{item[6]}, Soucet = {soucet}");
                //přepočet metry na stopa a formátování na dvě desetinná místa
                //string[] xx = [item[4], item[5], item[6], soucet.ToString("F2"), (soucet * 3.29).ToString("F2")];
                // Označen, počet vodičů, průřez, délka v metrech a délka ve stopách
                string[] xx = [item[4], item[5], item[6], soucet.ToString("F2")];
                Soucet.Add([.. xx]);
            }

            //Celkový kontrolní součet
            Soucet.Add([]);
            var celek = PoleData.Sum(x => double.TryParse(x[18], out double hodnota) ? hodnota : 0); // Převod textu na číslo a součet
            string[] xx1 = ["", "", "", celek.ToString("F2")];
            Soucet.Add([.. xx1]);

            //nová záložka
            //var ExcelApp = new ExcelApp();
            ExcelApp.GetSheet(SheetName);
            ExcelApp.Nadpis("A1:C1", "Označeni", Soucet.Count);
            
            ExcelApp.Nadpis("D1:D1", "Délka", Soucet.Count);
            ExcelApp.Nadpis("D2", "[m]");

            //ExcelApp.Nadpis("E1:E1", "Délka", Soucet.Count);
            //ExcelApp.Nadpis("E2", "[ft]");

            ExcelApp.KabelyToExcel(Soucet, 3);

            //xls.Cells[Soucet.Count + 1, 4].Formula = xxx může nastat chyba.
            ExcelApp.Xls.Cells[Soucet.Count + 1, 4].FormulaLocal = $"=SUMA(D3:D{Soucet.Count})"; // SUMAE{i}*500/480";
        }

    }
}
