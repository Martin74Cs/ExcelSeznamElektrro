// See https://aka.ms/new-console-template for more information
using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Exc = Microsoft.Office.Interop.Excel;

Item item = new Item();
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
if(File.Exists(CestaCsv)) File.Delete(CestaCsv);
data.ReadXml(CestaXML2);
Prevod.DataTabletoToCsv(data.Tables[0], CestaCsv);

string cesta = Path.Combine(BaseAdres, @"zarizeni.json");
var pokus = Soubory.LoadJsonEn<Item>(cesta);

Console.Write($"\nCelkem={pokus.Count()}");
Console.Write($"\n");
Vypis(pokus); 

ExcelApp Ex = new ExcelApp();

//Ex.ExcelSave(sheet, pokus.ToArray(), "Seznam zařízení");

string cestacelek = Path.Combine(BaseAdres, @"zarizeni_vse.xlsx");
var xlsc = ExcelApp.NovyExcelSablona(cestacelek);
var sheetc = ExcelApp.PridatNovyList(xlsc, "Seznam zažízení");
Ex.ExcelSave(sheetc, pokus.ToArray());
xlsc.Save();
//uzavření dokumentu bez uložení  
//xlsc.Close();
Ex.ExcelQuit(xlsc);

////Třdění jdenotlivých PS
//var Filtr = pokus.GroupBy(x => x._Item__eunit._Unit__pfx + " " + x._Item__eunit._Unit__num)
//    .Select(group => group).ToArray();

//foreach (var tr in Filtr)
//{
//    string cestap = Path.Combine(BaseAdres, @"zarizeni_" + tr.First()._Item__eunit._Unit__pfx + " " + tr.First()._Item__eunit._Unit__num + ".json");
//    string cestax = Path.ChangeExtension(cestap, ".xlsx");
//    //var xls = ExcelApp.VytvorNovyDokument();
//    var xls = ExcelApp.NovyExcelSablona(cestax);
//    var sheet = ExcelApp.PridatNovyList(xls, "Seznam zažízení");
//    Ex.ExcelSave(sheet, tr.ToArray());
//    xls.Save();
//    //uzavření dokumentu bez uložení  
//    //xlsc.Close();
//    Ex.ExcelQuit(xlsc);
//}

//Ex.ExcelSave(sheet, pokus.ToArray());
//xls.SaveAs2(Path.ChangeExtension(cesta, ".xlsx"));
//xls.Close();

//Ex.ExcelSave<Item>(sheet, pokus.ToArray(), "Pokus");
//Ex.ExcelSave(sheet, pokus.ToArray(), "Pokus");

void Vypis(List<Item> item)
{
    foreach (var i in item)
    {
        Console.WriteLine($"Tag={i._Item__tag}, Jmeno={i._Item__name}");
        if (i._Item__subitem.Count > 0)
            Vypis(i._Item__subitem);
    }
}

//Kopírování informací do revize 7
//Ele.Elektro();

//var Ele = new ElektroLoad();
////Vytvoření nového dokumentu
//Ele.NovyExcel();
