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
        public List<List<string>> LoadDataExcel(string cesta, int[] Sloupce, string Tabulka , int Radek)
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
            }
            else
            {
                Pole = new ExcelApp().ExelLoadTable(cesta, Tabulka, Radek, Sloupce);
                Pole.SaveJsonList(json);
            }
            return Pole;
        }

    }
}
