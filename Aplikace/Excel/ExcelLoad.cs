using Aplikace.Sdilene;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Excel
{
    public class ExcelLoad
    {

        /// <summary> Načtení dpkumentu Ecxel nebo Json do pole List<List<string>> z a vytvořejí JSON</summary>
        public static List<List<string>> LoadDataExcel(string cesta, int[] Sloupce, string Tabulka , int Radek, string[] TextPole)
        {
            Console.Write("\nProbíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            var Pole = new List<List<string>>();
            string Soubor = Path.GetFileName(cesta);
            string Adresar = Path.GetDirectoryName(cesta);
            string json = Path.Combine(Adresar, Path.ChangeExtension(Soubor, ".json"));
            if (File.Exists(json))
            {
                Pole = Soubory.LoadJsonList<List<string>>(json);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            }
            else
            {
                Pole = ExcelApp.ExelLoadTable(cesta, Tabulka, Radek, Sloupce, TextPole);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                if(Pole.Count>1) Pole.SaveJsonList(json);
            }
            Console.WriteLine($"načeno {Pole.Count} záznamů.");
            return Pole;
        }

        /// <summary> Načtení dpkumentu Ecxel nebo Json do pole List<List<string>> z a vytvořejí JSON</summary>
        public static List<Zarizeni> DataExcel(string cesta, string Tabulka, int Radek)
        {
            Console.WriteLine("Probíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            //var Pole = new List<List<string>>();
            var Pole = new List<Zarizeni>();
            string Soubor = Path.GetFileName(cesta);
            string Adresar = Path.GetDirectoryName(cesta);
            string json = Path.Combine(Adresar, Path.ChangeExtension(Soubor, ".json"));
            if (File.Exists(json))
            {
                Pole = Soubory.LoadJsonList<Zarizeni>(json);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            }
            else
            {
                if (!System.IO.File.Exists(cesta)) return [];

                var Xls = ExcelApp.DokumetExcel(cesta);
                if (Xls == null) return [];
                Console.WriteLine("Dokument excel - Otevřen");

                //Nastavení listu
                var Zal = ExcelApp.GetSheet(Xls, Tabulka);

                if (Zal == null) { Console.Write("\nChyba KONEC"); return []; }
                Console.WriteLine("Sheet=" + Zal.Name);
                Pole = ExcelApp.ExelTable(Zal, Tabulka, Radek);
                //Console.WriteLine("Zavřit dokument ");
                Xls.Close();

                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                if (Pole.Count > 1) Pole.SaveJsonList(json);
            }
            Console.WriteLine($"načeno {Pole.Count} záznamů.");
            return Pole;
        }


        /// <summary> Načtení dpkumentu Ecxel do pole Třídy z a vytvořejí JSON</summary>
        public static List<Zarizeni> LoadDataExcelTrida(string cesta, int[] Sloupce, string Tabulka , int Radek, string[] TextPole)
        {
            Console.Write("\nProbíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            var Pole = new List<Zarizeni>();
            string Soubor = Path.GetFileName(cesta);
            string Adresar = Path.GetDirectoryName(cesta);
            string json = Path.Combine(Adresar, Path.ChangeExtension(Soubor, ".json"));
            if (File.Exists(json))
            {
                Pole = Soubory.LoadJsonList<Zarizeni>(json);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            }
            else
            {
                Pole = ExcelApp.ExelLoadTableTrida(cesta, Tabulka, Radek, Sloupce, TextPole);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                Pole.SaveJsonList(json);
            }
            return Pole;
        }


    }
}
