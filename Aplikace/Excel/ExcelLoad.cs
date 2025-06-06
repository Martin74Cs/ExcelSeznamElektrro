using Aplikace.Sdilene;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Excel
{
    public class ExcelLoad
    {

        /// <summary> Načtení dokumentu Ecxel nebo Json do pole List<List<string>> z a vytvořejí JSON</summary>
        public static List<List<string>> LoadDataExcel(string cesta, int[] Sloupce, string Tabulka , int Radek)
        {
            Console.Write("\nProbíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            //var Pole = new List<List<string>>();
            //string Soubor = Path.GetFileName(cesta);
            //string Adresar = Path.GetDirectoryName(cesta);
            //string json = Path.Combine(Adresar, Path.ChangeExtension(Soubor, ".json"));
            string json = Path.ChangeExtension(cesta, ".json");
            if (File.Exists(json))
            {
                return Soubory.LoadJsonList<List<string>>(json);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            }
            else
            {
                var ExcelApp = new ExcelApp();
                var Pole = ExcelApp.ExelLoadTable(cesta, Tabulka, Radek, Sloupce);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                if(Pole.Count>1) Pole.SaveJsonList(json);
                Console.WriteLine($"načeno {Pole.Count} záznamů.");
                return Pole;
            }
        }

        /// <summary> Načtení dokumentu Ecxel nebo Json do pole List List string z a vytvořejí JSON</summary>
        public static List<Zarizeni> DataExcel(string cesta, string Tabulka, int Radek)
        {
            Console.WriteLine("Probíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            //var Pole = new List<List<string>>();
            //var Pole = new List<Zarizeni>();
            //string Soubor = Path.GetFileName(cesta);
            //string Adresar = Path.GetDirectoryName(cesta);

            if (!File.Exists(cesta)) return [];

            var ExcelApp = new ExcelApp(cesta);
            //ExcelApp.DokumetExcel(cesta);
            ExcelApp.GetSheet(Tabulka);
            if (ExcelApp.Xls == null) return [];
            
            if (ExcelApp.Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.WriteLine("Sheet=" + ExcelApp.Xls.Name);

            //Sloupce které se mají načíst z Excelu do názvů tříd. Myslím že třída musí existovat
            var dir = new Dictionary<int, string>() {
                {1, "Radek"     },
                {2, "Tag"       },
                {3, "Pocet"     },
                {4, "Popis"     },
                {11, "Menic"    },
                {10, "Prikon"   },
                {18, "BalenaJednotka"   },
            };

            var Pole = ExcelApp.ExelTable(Radek,Tabulka, dir);

            ExcelApp.ExcelQuit(cesta);
            //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            //if (Pole.Count > 1) Pole.SaveJsonList(Cesty.ElektroRozvaděčJson);
            Console.WriteLine($"načeno {Pole.Count} záznamů.");
            return Pole;
            
        }
        /// <summary> Načtení dokumentu Ecxel nebo Json do pole List Zarizeni z a vytvořejí JSON</summary>
        public static List<Zarizeni> DwgDataExcel(string cesta, string Tabulka, int Radek)
        {
            Console.WriteLine("Probíná hačítání dat ... ");
            if (!File.Exists(cesta)) return [];

            var ExcelApp = new ExcelApp(cesta);
            ExcelApp.GetSheet(Tabulka);
            if (ExcelApp.Xls == null) return [];
            
            if (ExcelApp.Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.WriteLine("Sheet=" + ExcelApp.Xls.Name);

            //Sloupce které se mají načíst z Excelu do názvů tříd. Myslím že třída musí existovat
            //DOPLNIT SLOUPCE PRO DWG
            var dir = new Dictionary<int, string>() {
                //{1, "Radek"   },
                {6, "Predmet"   },
                {7, "PID"       },
                //{3, "Pocet"   },
                {8, "Popis"     },
                {9, "Druh"      },
                {10, "Typ"      },
                {21, "Tag"      },
                {23, "TagStroj" },
                {24, "Menic"    },
                {26, "Prikon"   },
                {25, "Etapa"    },
                {27, "Patro"    },
                //{18, "BalenaJednotka"   },
            };

            var Pole = ExcelApp.ExelTable(Radek,Tabulka, dir);

            ExcelApp.ExcelQuit(cesta);
            Console.WriteLine($"Načeno {Pole.Count} záznamů.");
            return Pole;
            
        }

        /// <summary> Načtení dokumentu Ecxel do pole Třídy z a vytvořejí JSON</summary>
        public static List<Zarizeni> LoadDataExcelTrida(string cesta, int[] Sloupce, string Tabulka , int Radek, string[] TextPole)
        {
            if (!System.IO.File.Exists(cesta)) return [];
            Console.Write("\nProbíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            //var Pole = new List<Zarizeni>();
            string Soubor = Path.GetFileName(cesta);
            string Adresar = Path.GetDirectoryName(cesta) ?? Environment.SpecialFolder.MyDocuments.ToString();
            string json = Path.Combine(Adresar, Path.ChangeExtension(Soubor, ".json"));
            if (File.Exists(json))
            {
                return Soubory.LoadJsonList<Zarizeni>(json);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            }
            else
            {
                var ExcelApp = new ExcelApp();
                var Pole = ExcelApp.ExelLoadTableTrida(cesta, Tabulka, Radek, Sloupce, TextPole);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                Pole.SaveJsonList(json);
                return Pole;
            }
        }

        public static string Apid(int length = 9)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var random = new Random();
            //return new string(Enumerable.Repeat(chars, length)
            //    .Select(s => s[random.Next(s.Length)]).ToArray());
            return new string([.. Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)])]);
        }
    }
}
