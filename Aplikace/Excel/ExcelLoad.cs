using Aplikace.Sdilene;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Excel
{
    public class ExcelLoad
    {

        /// <summary> Načtení dpkumentu Ecxel do pole List<List<string>> z a vytvořejí JSON</summary>
        public List<List<string>> LoadDataExcel(string cesta, int[] Sloupce, string Tabulka , int Radek, string[] TextPole)
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
                Pole = new ExcelApp().ExelLoadTable(cesta, Tabulka, Radek, Sloupce, TextPole);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                Pole.SaveJsonList(json);
            }
            return Pole;
        }

    }
}
