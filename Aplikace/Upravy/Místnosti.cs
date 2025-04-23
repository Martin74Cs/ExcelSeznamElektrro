using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Upravy
{
    public class Místnosti
    {
        public static List<Mistnost> Vytvorit(string cesta, string objekt = "SO000")
        {
            //var cesta = Path.Combine(Cesty.BasePath, "revit", SO, "Místnosti.csv");
            //c:\a\LightChem\Elektro\revit\SO117\
            var Mistnost = Soubory.LoadFromCsv<Mistnost>(cesta);

            for (int i = 0; i < Mistnost.Count; i++)
            {
                Mistnost[i].Objekt = objekt;
            }
            Mistnost.SaveJsonList(Path.ChangeExtension(cesta, ".json"));
            Prevod.SaveToCsv(Mistnost ,Path.ChangeExtension(cesta, ".txt"));
            return Mistnost;
        } 
        public static void VytvoritSeznamy()
        {
            var cesta = Path.Combine(Cesty.BasePath, "revit", "SO117", "Výkaz místností.csv");
            var Misto =  Vytvorit(cesta, "SO117");

            cesta = Path.Combine(Cesty.BasePath, "revit", "SO119", "Výkaz místností.csv");
            var Misto2 = Vytvorit(cesta, "SO119");
            Misto.AddRange(Misto2);

            cesta = Path.Combine(Cesty.BasePath, "revit", "SO118", "Výkaz místností.csv");
            var Misto3 = Vytvorit(cesta, "SO118");
            Misto.AddRange(Misto3);

            var ExcelApp = new ExcelApp();
            ExcelApp.GetSheet("Místnosti");
            
            //Vytvoření nadpisů
            ExcelApp.Nadpisy<Mistnost>();

            //Vytvoření dat
            ExcelApp.ClassToExcel(Row: 2, Misto, Mistnost.Sloupce);
        }
    }
}
