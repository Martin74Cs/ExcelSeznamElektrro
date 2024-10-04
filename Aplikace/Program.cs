// See https://aka.ms/new-console-template for more information
using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Exc = Microsoft.Office.Interop.Excel;

Item item = new Item();
string BaseAdres = @"U:\Elektro\mcsato\Zakázky\Povrly.Med\";
string cestaPump = Path.Combine(BaseAdres, @"zarizeni.json"); 
if (File.Exists(cestaPump))
{
    var pokus = Soubory.LoadJsonEn<Item>(cestaPump);
    Console.Write($"\nCelkem={pokus.Count()}");
    Console.Write($"\n");
    Vypis(pokus);

    ExcelApp Ex = new ExcelApp();   
    var xls =  Ex.VytvorNovyDokument();
    var sheet = Ex.PridatNovyList(xls,"Seznam");
    Ex.ExcelSave(sheet, pokus.ToArray(), "Seznam zaziření");
    xls.SaveAs2(Path.ChangeExtension(cestaPump,".xlsx"));
    //xls.Close();
    xls.Application.Quit();
    
    Ex.ExcelSave<Item>(sheet, pokus.ToArray(), "Pokus");

    }

void Vypis(List<Item> item)
{
    foreach (var i in item)
    {
        Console.WriteLine($"Tag={i._Item__tag}, Patro={i._Item__name}");
        if (i._Item__subitem.Count > 0)
            Vypis(i._Item__subitem);
    }
}

//Kopírování informací do revize 7
//Ele.Elektro();

//var Ele = new ElektroLoad();
////Vytvoření nového dokumentu
//Ele.NovyExcel();