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
                Mistnost[i].Apid = ExcelLoad.Apid();
            }
            Mistnost.SaveJsonList(Path.ChangeExtension(cesta, ".json"));
            //Prevod.SaveToCsv(Mistnost ,Path.ChangeExtension(cesta, ".txt"));
            return Mistnost;
        }

        public static void VytvoritSeznamy()
        {
            //var Místnosti = Path.Combine(Cesty.BasePath, "Místnosti");
            //if (!Directory.Exists(Místnosti)) Directory.CreateDirectory(Místnosti);          
            //var Revit = Path.Combine(Místnosti, "revit");
            //if (!Directory.Exists(Revit)) Directory.CreateDirectory(Revit);

            var Slabo = new List<Slaboproudy>();

            //Vstupní data
            var cesta = Path.Combine(Cesty.Revit, "SO117", "Výkaz místností.csv");
            var Misto = Vytvorit(cesta, "SO117");
            var parents = Misto.Select(Slaboproudy.CopyToParent).ToList();
            Slabo.AddRange(parents);

            cesta = Path.Combine(Cesty.Revit, "SO119", "Výkaz místností.csv");
            var Misto2 = Vytvorit(cesta, "SO119");
            Misto.AddRange(Misto2);
            var parents2 = Misto2.Select(x => Slaboproudy.CopyToParent(x)).ToList();
            Slabo.AddRange(parents2);

            cesta = Path.Combine(Cesty.Revit, "SO118", "Výkaz místností.csv");
            var Misto3 = Vytvorit(cesta, "SO118");
            Misto.AddRange(Misto3);
            var parents3 = Misto3.Select(x => Slaboproudy.CopyToParent(x)).ToList();
            Slabo.AddRange(parents3);

            //Hlavní soubor
            //string cestaXLs = Path.Combine(Místnost, "Místnosti.celek.xlsx");
            Console.WriteLine(Cesty.MistnostiXLs);

            //Soubory pro upravení
            Slabo.SaveJsonList(Path.ChangeExtension(Cesty.MistnostiXLs, ".json"));
            Slabo.SaveToCsv(Path.ChangeExtension(Cesty.MistnostiXLs, ".csv"));

            //Vyvořit nebo otevřít excel
            var ExcelApp = new ExcelApp(Cesty.MistnostiXLs);
            ExcelApp.GetSheet("Místnosti");
            
            //Vytvoření nadpisů
            //ExcelApp.Nadpisy<Mistnost>(Mistnost.Sloupce);
            ExcelApp.Nadpisy(Slaboproudy.SloupceSpojit);

            //Vytvoření dat
            ExcelApp.ClassToExcel(Row: 2, Slabo, Slaboproudy.SloupceSpojit);
            //Uložení a ukončení
            ExcelApp.ExcelQuit(Cesty.MistnostiXLs);
        }
    }
}
