using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Excel
{
    //public class GetData
    //{
    //    /// <summary> Načtení dat z Excelu Formátování a text </summary>
    //    public void Data(string Cesta)
    //    {
    //        if (string.IsNullOrEmpty(Cesta))
    //        {
    //            Console.WriteLine("Cesta k souboru je prázdná.");
    //            return;
    //        }
    //        var excelApp = new Exc.Application();
    //        var workbook = excelApp.Workbooks.Open(Cesta);
    //        var worksheet = (Exc.Worksheet)workbook.Sheets["List1"]; // Změň název listu

    //        Console.WriteLine("Sloučené buňky:");

    //        Exc.Range slouceneBunky = worksheet.UsedRange.MergeCells ? worksheet.UsedRange.MergeArea : null;
    //        Exc.Range mergedCells = worksheet.UsedRange;

    //        foreach (Exc.Range area in mergedCells.Cells)
    //        {
    //            if ((bool)(area.MergeCells))
    //            {
    //                Exc.Range merged = area.MergeArea;

    //                string rozsah = merged.get_Address();
    //                string hodnota = Convert.ToString(merged.Cells[1, 1].Value);

    //                Console.WriteLine($"Rozsah: {rozsah}, Hodnota: {hodnota}");
    //            }
    //        }

    //        workbook.Close(false);
    //        excelApp.Quit();
    //    }
    //}
}
