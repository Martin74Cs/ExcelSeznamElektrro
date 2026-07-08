﻿using Knihovna;
using Knihovna.Excel;
using Knihovna.Sdilene;
using System.Data;

namespace Aplikace.Upravy
{
    public class Povrly
    {
        /// <summary> Převody souborů z JSON do XML a CSV </summary>
        public static void Hlavni()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr souboru 'zarizeni.json'...");
            Console.ResetColor();
            string? cesta1 = Soubory.ShowOpenFileDialog("JSON soubory (*.json)|*.json");
            if (string.IsNullOrEmpty(cesta1) || !File.Exists(cesta1))
            {
                Console.WriteLine("Výběr souboru byl stornován nebo soubor neexistuje.");
                return;
            }

            string? BaseAdres = Path.GetDirectoryName(cesta1);
            if (string.IsNullOrEmpty(BaseAdres)) return;

            string jsonString = System.IO.File.ReadAllText(cesta1);
            //převod souboru
            string XML2 = Prevod.JsonToXml(jsonString);
            string CestaXML2 = Path.Combine(BaseAdres, @"zarizeni2.xml");
            File.WriteAllText(CestaXML2, XML2);

            string XML = Prevod.JsonToXmlAI(jsonString);
            string CestaXML = Path.Combine(BaseAdres, @"zarizeni.xml");
            File.WriteAllText(CestaXML, XML);

            string CestaCsv = Path.Combine(BaseAdres, @"zarizeni.csv");
            //přeovd a save do Csv
            var data = new DataSet();
            //načtení z ulolženého souboru
            if (File.Exists(CestaCsv)) File.Delete(CestaCsv);
            data.ReadXml(CestaXML2);
            Prevod.DataTabletoToCsv(data.Tables[0], CestaCsv);

            var pokus = Soubory.LoadJsonEn<Knihovna.Tridy.Item>(cesta1);

            Console.Write($"\nCelkem={pokus.Count}");
            Console.Write($"\n");
            Vypis(pokus);


            string cestacelek = Path.Combine(BaseAdres, @"zarizeni_vse.xlsx");
            //ExcelApp
            var ExcelApp = new ExcelApp(cestacelek);
            ExcelApp.GetSheet("Seznam zažízení");
            ExcelApp.ExcelSave([.. pokus]);
            ExcelApp.Doc.Save();
            //uzavření dokumentu bez uložení  
            ExcelApp.ExcelQuit(cestacelek);
        }
        static void Vypis(List<Knihovna.Tridy.Item> item)
        {
            foreach (var i in item)
            {
                Console.WriteLine($"Tag={i.Tag}, Jmeno={i.Name}");
                if (i.Subitem.Count > 0)
                    Vypis(i.Subitem);
            }
        }
    }
}

