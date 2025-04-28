using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Upravy
{
    public class Povrly
    {
        /// <summary> Převody souborů z JSON do XML a CSV </summary>
        public static void Hlavni()
        {
            //Item item = new Item();
            string BaseAdres = @"U:\Elektro\mcsato\Zakázky\Povrly.Med\";
            string cesta1 = Path.Combine(BaseAdres, @"zarizeni.json");
            if (!File.Exists(cesta1)) return;

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
            //Prevod.JsonToCsv(jsonString, CestaCsv);
            var data = new DataSet();
            //načtení z ulolženého souboru
            if (File.Exists(CestaCsv)) File.Delete(CestaCsv);
            data.ReadXml(CestaXML2);
            Prevod.DataTabletoToCsv(data.Tables[0], CestaCsv);

            string cesta = Path.Combine(BaseAdres, @"zarizeni.json");
            var pokus = Soubory.LoadJsonEn<Item>(cesta);

            Console.Write($"\nCelkem={pokus.Count}");
            Console.Write($"\n");
            Vypis(pokus);

            //Ex.ExcelSave(sheet, pokus.ToArray(), "Seznam zařízení");

            string cestacelek = Path.Combine(BaseAdres, @"zarizeni_vse.xlsx");
            var ExcelApp = new ExcelApp(cestacelek);
            //ExcelApp.NovyExcelSablona(cestacelek);
            //Worksheet Xls = Doc.Worksheets[1];
            ExcelApp.GetSheet("Seznam zažízení");
            ExcelApp.ExcelSave([.. pokus]);
            ExcelApp.Doc.Save();
            //uzavření dokumentu bez uložení  
            //xlsc.Close();
            ExcelApp.ExcelQuit(cestacelek);
        }
        static void Vypis(List<Item> item)
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
