// See https://aka.ms/new-console-template for more information
using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using Aplikace.Upravy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using Exc = Microsoft.Office.Interop.Excel;

//Soubory.KillExcel();

//var Ele = new ElektroLoad();
//Vytvoření nového dokumentu
//ElektroLoad.NovyExcel();

//Lithtchem
//var LigthChem = new LigthChem();
//LigthChem.Hlavni();

//Mistnoti
//Místnosti.VytvoritSeznamy();

//Rozvaděč
LigthChem.Rozvadec();


////Třdění jdenotlivých PS
//var Filtr = pokus.GroupBy(x => x.cunit.pfx + " " + x.cunit.num)
//    .Select(group => group).ToArray();

//foreach (var tr in Filtr)
//{
//    string cestap = Path.Combine(BaseAdres, @"zarizeni_" + tr.First().cunit.pfx + " " + tr.First().cunit.num + ".json");
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

//Kopírování informací do revize 7
//Ele.Elektro();

