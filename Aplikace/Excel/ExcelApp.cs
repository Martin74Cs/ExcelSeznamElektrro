using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Excel
{
    public class ExcelApp
    {
        /// <summary> Vytvoření nového Excel dokumentu </summary>
        public Exc.Workbook? VytvorNovyDokument()
        {
            Exc.Application? App = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
            if (App == null) return null;
            App.Visible = true;

            // Vytvoření nového sešitu
            Exc.Workbook NovyDokument = App.Workbooks.Add();
            Console.Write("\nNový Excel dokument byl vytvořen.");
            return NovyDokument;
        }

        /// <summary> Přidání nového listu do Excelového dokumentu </summary>
        public Exc.Worksheet PridatNovyList(Exc.Workbook Dokument, string NazevListu)
        {
            Exc.Worksheet? xls = new ExcelApp().GetSheet(Dokument, NazevListu);
            if(xls == null)
                // Přidání nového listu na konec sešitu
                xls = Dokument.Sheets.Add(After: Dokument.Sheets[Dokument.Sheets.Count]);
            xls.Name = NazevListu;
            return xls;
        }

        /// <summary> nastavení listu dle jeho jména</summary>
        public Exc.Worksheet? GetSheet(Exc.Workbook Dokument, string Nazev)
        {
            foreach (Exc.Worksheet item in Dokument.Sheets)
            {
                if (item.Name == Nazev)
                    return item;
            }
            return null;
        }

        /// <summary>Nový dokument v exelu</summary>
        public Exc.Workbook? DokumetExcel(string Cesta)
        {
            //Exc.Application App = AplikaceExcel();
            Exc.Application? App = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
            if (App == null) return null;
            App.Visible = true;

            Console.Write("\nKontrolaOtevenehoNeboOtevreniSobroruExel - OK");
            return KontrolaOtevenehoNeboOtevreniSobroruExel(App, Cesta);
        }

        /// <summary>Kontrola otevřeného souboru v Excel</summary>
        public Exc.Workbook KontrolaOtevenehoNeboOtevreniSobroruExel(Exc.Application ExApp, string Cesta)
        {
            Console.Write("\nMetoda Kontrola Oteveneho Nebo Otevreni Sobroru Exel");
            Console.Write("\nCesta" + Cesta.ToLowerInvariant());
            foreach (Exc.Workbook item in ExApp.Workbooks)
            {
                Console.Write("\nName=" + item.Name);
                if (item.Name == System.IO.Path.GetFileName(Cesta.ToLowerInvariant()))
                    return item;
            }
            Console.Write("\nSoubor není otevřen kontrola ");
            //return null;
            //nefunuguje otevření souboru
            return ExApp.Workbooks.Open(Cesta.ToLowerInvariant());
        }

        //public Exc.Application AplikaceExcel()
        //{
        //    try
        //    {
        //        if (ExcelKontrolaInstalace() == false)
        //        {
        //            Console.Write("\nExcelKontrolaInstalace");
        //            return new Exc.Application();
        //        }
        //        Console.Write("\nMarshal.GetActiveObject");
        //        //return System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Exc.Application;

        //        //vytvoříte instanci Excelu, pokud již neběží, a pokud běží, připojíte se k aktivní instanci.
        //        dynamic excelApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        //        return excelApp;
        //    }
        //    catch (System.Runtime.InteropServices.COMException)
        //    {
        //        return new Exc.Application();
        //    }
        //}

        /// <summary> uložení dat do excel podle kriterii </summary>
        public List<List<string>> ExelLoadTable(string cesta, string zalozka, int Radek, int[] CteniSloupcu)
        {
            if (!System.IO.File.Exists(cesta)) return [];

            var Xls = DokumetExcel(cesta);
            if (Xls == null) return new List<List<string>>();
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            Exc.Worksheet? Zal = GetSheet(Xls, zalozka);
            if (Zal == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.Write("\nSheet=" + Zal.Name);

            var cteniPole = new List<string>();
            var Pole = new List<List<string>>();
            Console.Write("\nZal.Rows.Count=" + Zal.Rows.Count);
            for (int i = Radek; i < Zal.Rows.Count; i++)
            {
                cteniPole = new List<string>();
                foreach (var item in CteniSloupcu)
                {
                    //Čtení buňky
                    Exc.Range Pok = Zal.Cells[i, item];
                    object cteni = Pok.Value;

                    string xxx = Convert.ToString(cteni);
                    if (!string.IsNullOrEmpty(xxx))
                        cteniPole.Add(xxx);
                    else
                        cteniPole.Add("");
                }

                if (!string.IsNullOrEmpty(cteniPole[1]) && cteniPole[1] != "0")
                {
                    Pole.Add(cteniPole);
                    Console.Write("\nRadek=" + i.ToString());
                }

                if (i > 500) break;
            }
            Console.Write("\nUkončení Excel");
            //Xls.Save();
            //Console.Write("\nSave OK");
            ExcelQuit(Xls);
            Console.Write("\nUkončení Excel");
            return Pole;
        }


        /// <summary> uložení dat do excel podle kdyterii </summary>
        public void ExcelSaveJeden(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
        {
            if (!System.IO.File.Exists(cesta)) return;

            var Xls = DokumetExcel(cesta);
            if (Xls == null) return;
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            Exc.Worksheet? Zal = GetSheet(Xls, zalozka);
            if (Zal == null) { Console.Write("\nChyba KONEC");  return; }
            //Exc.Worksheet Zal = Xls.Worksheets[zalozka];
            Console.Write("\nSheet=" + Zal.Name);

            //Čtení listu excel
            for (int i = 7; i < Zal.Rows.Count; i++)
            {
                //Čtení radků excel
                var cteniPole = new List<string>();
                foreach (var item in SloupceCteni)
                {
                    //Čtení buňky Tag
                    Exc.Range Pok = Zal.Cells[i, item];
                    string xxx = Convert.ToString(Pok.Value);
                    if (!string.IsNullOrEmpty(xxx))
                        cteniPole.Add(xxx);
                }

                //Hledání shody Vstupu s načteným řádkem Hledání v první shode
                var Shoda = Vstup.FirstOrDefault(x => x.FirstOrDefault() == cteniPole.FirstOrDefault());

                //Pokud byla nalezeny schoda radku s polem vstupu
                if (Shoda != null)
                {
                     Console.Write("\nShoda buňky " + i + " = " + Shoda.First());

                    //zapis buňky
                    Exc.Range Zapis = Zal.Cells[i, SloupceZapisu.First()];
                    Zapis.Value = Shoda.First();

                    //Posledni
                    Zapis = Zal.Cells[i, SloupceZapisu.Last()];
                    Zapis.Value = Shoda[8] + " " + Shoda.Last();
                }
                else
                {
                    //nebyla shoda
                    //zapis buňky
                    Exc.Range Zapis = Zal.Cells[i, SloupceZapisu.First()];
                    Zapis.Value = "Nenalezeno";
                }

                { Console.Write("\nShoda buňky " + i); }
                if (i > 500) break;
            }
            Console.Write("\nUkončení Excel");
            Xls.Save();

            Console.Write("\nSave OK");
            //Xls.Close();
            //ed.WriteMessage("\nClose OK");
            ExcelQuit(Xls);
            Console.Write("\nUkončení Excel");
            return;
        }

        /// <summary> uložení dat do excel podle kdyterii </summary>
        public void ExcelSaveSloupec(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
        {
            string cesta1 = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
            var PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = new ExcelLoad().LoadDataExcel(cesta1, PouzitProTabulku, "Motory500V", 2);

            if (!System.IO.File.Exists(cesta)) return;

            var Xls = DokumetExcel(cesta);
            if (Xls == null) return;
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            Exc.Worksheet? Zal = GetSheet(Xls, zalozka);
            if (Zal == null) return;
            //Exc.Worksheet Zal = Xls.Worksheets[zalozka];
            Console.Write("\nSheet=" + Zal.Name);

            //Čtení listu excel
            for (int i = 7; i < Zal.Rows.Count; i++)
            {
                //Čtení radků excel
                var cteniPole = new List<string>();
                foreach (var item in SloupceCteni)
                {
                    //Čtení buňky
                    Exc.Range Pok = Zal.Cells[i, item];
                    object cteni = Pok.Value;

                    string xxx = Convert.ToString(cteni);
                    if (!string.IsNullOrEmpty(xxx))
                        cteniPole.Add(xxx);
                }

                //Hledání shody Vstupu s načteným řádkem Hledání v první shode
                var Shoda = Vstup.FirstOrDefault(x => x.FirstOrDefault() == cteniPole.FirstOrDefault());

                //Pokud byla nalezeny schoda radku s polem vstupu

                if (Shoda != null)
                {
                    Console.Write("\nShoda buňky " + i + " = " + Shoda.First()); 

                    //hledni proudu z tabulky Motory500V
                    if (double.TryParse(cteniPole[1], out double Prikon))
                    {
                        var Informace = Motory500.FirstOrDefault(x => Convert.ToDouble(x[0]) == Prikon)?[1]; //.ToArray(); 
                        if (double.TryParse(Informace, out double Proud))
                        {
                            Exc.Range Zapis1 = Zal.Cells[i, SloupceZapisu.First()];
                            Zapis1.Value = Proud;
                        }
                    }

                    ////zapis proud
                    //Exc.Range Zapis = Zal.Cells[i, SloupceZapisu.First()];
                    //if (double.TryParse(Shoda[3], out double cislo))
                    //    Zapis.Value = cislo;
                    //else
                    //    Zapis.Value = "";

                    //Rozvaděč
                    var Zapis = Zal.Cells[i, SloupceZapisu[1]];
                    Zapis.Value = Shoda[8];

                    //Rozvaděč
                    Zapis = Zal.Cells[i, SloupceZapisu[2]];
                    Zapis.Value = Shoda[9];

                    //zapis delka
                    Zapis = Zal.Cells[i, SloupceZapisu[3]];
                    if (double.TryParse(Shoda[4], out double delka))
                        Zapis.Value = delka;
                    else
                        Zapis.Value = Shoda[4].ToString();

                    //zapis AWG
                    Zapis = Zal.Cells[i, SloupceZapisu[5]];
                    if (double.TryParse(Shoda[5], out double AWG))
                        Zapis.Value = AWG;
                    else
                        Zapis.Value = Shoda[5].ToString();

                    //zapis mm2
                    Zapis = Zal.Cells[i, SloupceZapisu[4]];
                    if (double.TryParse(Shoda[10], out double mm2))
                        Zapis.Value = mm2;
                    else
                        Zapis.Value = "";
                }
                else
                {
                    //nebyla shoda
                    //zapis buňky
                    Exc.Range Zapis = Zal.Cells[i, SloupceZapisu.First()];
                    Zapis.Value = "Nenalezeno";
                }

                { Console.Write("\nShoda buňky " + i); }
                if (i > 500) break;
            }
            Console.Write("\nUkončení Excel");
            Xls.Save();

            Console.Write("\nSave OK");
            //Xls.Close();
            //ed.WriteMessage("\nClose OK");
            ExcelQuit(Xls);
            Console.Write("\nUkončení Excel");
            return;
        }

        public void NadpisSet(Worksheet ListExcel)
        {
            //Podtržení nadpisů

            // Výběr konkrétní oblasti buněk, např. A1:C3
            Exc.Range range = ListExcel.Range["A1", "M1"];

            // Definování rozsahu pomocí čísel řádků a sloupců (např. A1:C3)
            Exc.Range range1 = ListExcel.Range[
                ListExcel.Cells[2, 1],  // A1 (1. řádek, 1. sloupec)
                ListExcel.Cells[2, 12]   // C3 (3. řádek, 3. sloupec)
            ];

            // Nastavení okrajů kolem buněk
            // LineStyle: Může být xlContinuous, xlDash, xlDot a další styly čar.
            range.Borders[Exc.XlBordersIndex.xlEdgeLeft].LineStyle = Exc.XlLineStyle.xlContinuous;
            range.Borders[Exc.XlBordersIndex.xlEdgeRight].LineStyle = Exc.XlLineStyle.xlContinuous;
            range.Borders[Exc.XlBordersIndex.xlEdgeTop].LineStyle = Exc.XlLineStyle.xlContinuous;
            range.Borders[Exc.XlBordersIndex.xlEdgeBottom].LineStyle = Exc.XlLineStyle.xlContinuous;

            // Další možnosti nastavení tloušťky a barvy okrajů
            range.Borders.LineStyle = Exc.XlLineStyle.xlContinuous;

            // Weight: Určuje tloušťku čáry(xlThin, xlMedium, xlThick).
            range.Borders.Weight = Exc.XlBorderWeight.xlMedium;  // nebo xlMedium, xlThick - tlustá

            //Color: Převádí barvu z knihovny System.Drawing.Color na formát použitelný v Excelu.
            range.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black); // nastavení barvy čar

            // Nastavení barvy textu (např. červená)
            range.Font.Color = ColorTranslator.ToOle(Color.Red);
            //range.Font.Color = ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));  // Červená barva

            //Tučné písmo
            range.Font.Bold = true;
            //range.Font.Italic = true;

            //Velikost písma
            range.Font.Size = 14;
            range.Font.FontStyle = "Arial";

            //Vycentruje text vodorovně.
            range.HorizontalAlignment = Exc.XlHAlign.xlHAlignCenter;

            //Vycentruje text svisle
            range.VerticalAlignment = Exc.XlVAlign.xlVAlignCenter;

            //Orientace textu 
            range.Orientation = 90;

            // Nastavení barvy buňky (pozadí) (např. světle modrá)
            range.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
            //range.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(173, 216, 230));  // Světle modrá

            // Automatické přizpůsobení šířky sloupce (např. pro sloupec A)
            //Xls.Columns[1].AutoFit();

            // Automatické přizpůsobení výšky řádku (např. pro řádek 1)
            //Xls.Rows[1].AutoFit();

            //range 
            range.Rows["1"].AutoFit();
            //range.Columns["A:M"].AutoFit();
        }


        public void Nadpis(Worksheet Xls)
        {
            Xls.Range["A1"].Value = "Equipment\nnumber";

            Xls.Range["B1"].Value = "Equipment name";

            Xls.Range["C1"].Value = "Power(electric)\n(EU Units)";
            Xls.Range["C1"].Value = "[kW]";

            Xls.Range["D1"].Value = "Package unit Power(electric)\n(EU Units)";
            
            Xls.Range["E1"].Value = "PROUD Z TAB. PRO 500V";
            Xls.Range["E2"].Value = "[A]";

            Xls.Range["F1"].Value = "Power(electric)\n(US Units)";
            Xls.Range["F2"].Value = "[HP]";

            Xls.Range["G1"].Value = "CURRENT FOR 480V";
            Xls.Range["G2"].Value = "[A]";

            Xls.Range["H1"].Value = "CABLE LENGHT\n[m]";
            Xls.Range["H2"].Value = "[m]";

            Xls.Range["I1"].Value = "CABLE LENGHT\n[ft]";
            Xls.Range["I2"].Value = "[ft]";

            Xls.Range["J1"].Value = "COPPER CABLE SIZE\n(EU Units)[mm2]";
            Xls.Range["J2"].Value = "[mm2]";

            Xls.Range["K1"].Value = "COPPER CABLE SIZE\n(US Units)[ft]";
            Xls.Range["K2"].Value = "[ft]";

            Xls.Range["L1"].Value = "DISTRIBUTOR\nEA/MCC";

            Xls.Range["M1"].Value = "DISTRIBUTOR\nNUMBER";
            
            // Povolení zalamování textu, aby nový řádek byl viditelný
            Xls.Range["A1:M1"].WrapText = true;
        }

        /// <summary> uložení dat do excel podle kryterii </summary>
        public void ExcelSaveList(Worksheet Xls, List<List<string>> Vstup)
        {
            int col = 0; int row = 2;
            Nadpis(Xls);
            NadpisSet(Xls);

            //Čtení listu excel
            foreach (var radek in Vstup)
            {
                //Čtení radků excel
                var cteniPole = new List<string>();
                if (radek[2] != "" && radek[2] != "0")
                { 
                    col=1; row++;
                    foreach (var item in radek)
                    {
                        //zapis qwe
                        var Zapis = Xls.Cells[row, col++];
                        if (double.TryParse(item, out double cislo))
                            Zapis.Value = cislo;
                        else 
                        {
                            if (item == "PU")
                            {
                                Zapis = Xls.Cells[row, col - 2];
                                Zapis.Value = item;
                            }
                            else
                                Zapis.Value = item;
                        }
                    }
                }
            }
            return;
        }

        /// <summary> uložení dat do excel podle kdyterii </summary>
        public void ExcelSaveProud(Worksheet ListExcel, List<List<string>> Vstup)
        {

            //Čtení listu excel
            for (int i = 3; i < ListExcel.Rows.Count; i++)
            {
                //Čtení kW
                Exc.Range Pok = ListExcel.Cells[i, 3];
                object cteni = Pok.Value;

                string xxx = Convert.ToString(cteni);
                if (double.TryParse(xxx, out double cislo))
                {
                    //hledni proudu z tabulky Motory500V
                    var Informace = Vstup.FirstOrDefault(x => Convert.ToDouble(x[0]) == cislo)?[1]; //.ToArray(); 
                    if (double.TryParse(Informace, out double Proud))
                    {
                        Exc.Range Zapis1 = ListExcel.Cells[i, 5];
                        Zapis1.Value = Proud;
                    }
                }

                // Dynamický vzorec (např. sčítání hodnot v buňkách A a B na daném řádku)
                //string formula = $"=A{row}+B{row}";
                //string formula = $"=Cells({i}, 3)+Cells({3}, 2)";
                //string formula = $"=Cells({i}, 3)*1,34102";
                //ListExcel.Cells[i, 6].Formula = formula;

                // Dynamický vzorec pomocí Excelové notace (např. C pro sloupec 3)
                string formula = $"=C{i}*1.34102";  // C{i} odkazuje na buňku ve sloupci C (3) a řádku i

                // Vložení vzorce do sloupce 6 (odpovídá sloupci F)
                ListExcel.Cells[i, 6].Formula = formula;


                formula = $"=E{i}*500/480";
                ListExcel.Cells[i, 7].Formula = formula;

                formula = $"=H{i}*3.29";
                ListExcel.Cells[i, 9].Formula = formula;

                if (cteni == null)
                    break;
            }
            return;
        }




        /// <summary> Ze zadaného listu Exel vytvoř DataTable - podle zvolených sloupců </summary>
        public System.Data.DataTable GetTable(Exc.Worksheet oSheet, int rowNadpis, int[] sloupec)
        {
            //ed.WriteMessage("\nZačala metoda GetTable");
            //ed.WriteMessage("\nNadpis=" + rowNadpis + ", sloupec=" + sloupec.Length + ", Name=" + oSheet.Name + ", Rows=" + oSheet.Rows.Count);

            var Table = new System.Data.DataTable("Tabulka");

            // Načtěte konkrétní řádek
            Exc.Range rowRange = oSheet.Rows[rowNadpis];
            //ed.WriteMessage("\nVelikost Sheet.Rows " + rowRange.Columns.Count); //vysledek je 16384

            Exc.Range range = oSheet.UsedRange;
            int usedRows = range.Rows.Count;
            int usedCols = range.Columns.Count;
            //ed.WriteMessage("\nVelikost Table " + usedRows + ", " + usedCols);

            int colPomoc = 0;
            //Vytvoření nadpisů
            foreach (var i in sloupec)
            {
                //ed.WriteMessage("\nSloupec " + i);
                Exc.Range cell = rowRange.Cells[i];
                //ed.WriteMessage("\nFunguje");
                string cellValue = cell.Value?.ToString().Trim();
                Console.Write("\nRadek=" + rowNadpis + ", Sloupec=" + i + ", nadpis=" + cellValue);
                Table.Columns.Add(cellValue ?? i.ToString(), typeof(string));
                //Table.Columns.Add(i.ToString(), typeof(string));
            }
            Console.Write("\ninfo" + usedRows + ", " + usedCols + ", " + colPomoc);

            int t = 0;
            for (int row = rowNadpis + 1; row < usedRows; row++)
            {
                var Pole = new List<string>();
                //seznam sloupců ze zadání

                //DataRow rada;
                var rada = Table.NewRow();
                int colpomoc = 0;
                string text = string.Empty;
                foreach (var col in sloupec)
                {
                    //čtení buňky
                    Exc.Range Pok = range.Cells[row, col];
                    var cteni = Convert.ToString(Pok.Value);
                    //if(string.IsNullOrEmpty(cteni))                  
                    Pole.Add(cteni);
                    text += cteni;
                    rada[colpomoc++] = cteni;
                    //ed.WriteMessage("\ncteni " + cteni + "Pocet="  + Pole.Count);
                    Console.Write("\ncteni " + cteni);
                }
                Table.Rows.Add(rada);

                //Kontrola konce
                Console.Write("\nDelka" + text.Length);
                if (text.Length < 4) return Table;
                if (t > 1000) return Table;
            }
            return Table;
        }

        // <summary>Kontrola instalovaného Excelu false - Aplikace Exel není instalována</summary>
        public bool ExcelKontrolaInstalace()
        {
            if (Type.GetTypeFromProgID("Excel.Application") != null)
                return true;
            return true;
        }

        //ukončení worksheet
        public bool ExcelQuit(Exc.Workbook work)
        {
            // Ukončení aplikace Excel
            work.Application.Quit();
            return true;
        }

        public void ExcelSaveKabel(Worksheet ListExcel, List<List<string>> Vstup)
        {

            //Čtení listu excel
            for (int i = 2; i < ListExcel.Rows.Count; i++)
            {
                //Čtení nazvu
                Exc.Range Pok = ListExcel.Cells[i, 1];
                object cteni = Pok.Value;

                string xxx = Convert.ToString(cteni);

                //hledni proudu z tabulky delka
                var Informace = Vstup.FirstOrDefault(x => x[0] == xxx); //.ToArray(); 

                //delka
                if (double.TryParse(Informace?[4], out double delka))
                {
                    Exc.Range Zapis1 = ListExcel.Cells[i, 8];
                    Zapis1.Value = delka;
                }

                //awg
                Exc.Range Zapis = ListExcel.Cells[i, 11];
                Zapis.Value = Informace?[5];

                //mm2
                if (double.TryParse(Informace?[10], out double mm2))
                {
                    Exc.Range Zapis1 = ListExcel.Cells[i, 10];
                    Zapis1.Value = mm2;
                }

                if (cteni == null && i > 100)
                    break;
            }
        }

        public void ExcelSaveRozvadec(Worksheet ListExcel, List<List<string>> Vstup)
        {
            //Čtení listu excel
            for (int i = 2; i < ListExcel.Rows.Count; i++)
            {
                //Čtení nazvu
                Exc.Range Pok = ListExcel.Cells[i, 1];
                object cteni = Pok.Value;
                string xxx = Convert.ToString(cteni);

                //hledni proudu z tabulky delka
                var Informace = Vstup.FirstOrDefault(x => x[0] == xxx); //.ToArray(); 

                //mcc
                Exc.Range Zapis = ListExcel.Cells[i, 12];
                Zapis.Value = Informace?[8];

                //mcc
                if (double.TryParse(Informace?[9], out double cislo))
                {
                    Exc.Range Zapis1 = ListExcel.Cells[i, 13];
                    Zapis1.Value = cislo;
                }

                if (cteni == null && i > 100)
                    break;
            }
        }

        public List<List<string>> ExcelLoadWorksheet(Worksheet xls, int[] pouzitProTabulku)
        {
            var Data = new List<List<string>>();
            string Cteni = "";
            //Čtení listu excel
            for (int i = 3; i < xls.Rows.Count; i++)
            {
                var Radek = new List<string>();
                foreach (var item in pouzitProTabulku)
                {
                    //zapis qwe
                    var Zapis = xls.Cells[i, item];
                    Cteni = Convert.ToString(Zapis.Value);
                    Radek.Add(Cteni);
                }
                Data.Add(Radek);

                if (Cteni == null && i > 100)
                    break;
            }
            return Data;
        }

        public void ExcelSaveTable(Worksheet xls, List<List<string>> data, int Radek)
        {
            Radek--;
            foreach (var radek in data)
            {
                Radek++;int j = 1;
                foreach (var item in radek)
                {
                    Exc.Range Zapis1 = xls.Cells[Radek, j++];
                    if (double.TryParse(item, out double cislo))
                    {
                        Zapis1.Value = cislo;
                    }
                    else 
                    {
                        Zapis1.Value = item;
                    }
                }
            }
        }

        public void ExcelSaveNadpis(Worksheet xls)
        {
            Nadpis(xls, "A1:C1", "Označeni");
            Nadpis(xls, "D1:G1", "Kabel");
            Nadpis(xls, "H1:H1", "Zařízení");
            Nadpis(xls, "I1:K1", "Odkud");
            Nadpis(xls, "L1:N1", "Kam");
            Nadpis(xls, "O1:P1", "Delka");

            xls.Range["E2"].Value = "[mm2]";
            xls.Range["F2"].Value = "[AWG]";
            xls.Range["O2"].Value = "[m]";
            xls.Range["P2"].Value = "[ft]";
        }

        public void Nadpis(Worksheet xls, string pole, string Text)
        {
            // Sloučení buněk od A1 do C1
            var rada = xls.Range[pole];
            rada.Merge();
            rada.Value = Text;

            //zarovnání
            rada.HorizontalAlignment = Exc.XlHAlign.xlHAlignCenter;
            rada.VerticalAlignment = Exc.XlVAlign.xlVAlignCenter;

            // Další možnosti nastavení tloušťky a barvy okrajů
            rada.Borders.LineStyle = Exc.XlLineStyle.xlContinuous;

            // Weight: Určuje tloušťku čáry(xlThin, xlMedium, xlThick).
            rada.Borders.Weight = Exc.XlBorderWeight.xlMedium;  // nebo xlMedium, xlThick - tlusté

            //Color: Převádí barvu z knihovny System.Drawing.Color na formát použitelný v Excelu.
            rada.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black); // nastavení barvy čar

            // Nastavení barvy textu (např. červená)
            rada.Font.Color = ColorTranslator.ToOle(Color.Red);
            //range.Font.Color = ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));  // Červená barva

            //Tučné písmo
            rada.Font.Bold = true;
            //range.Font.Italic = true;

            //Velikost písma
            rada.Font.Size = 14;

            //Styl pisma
            rada.Font.FontStyle = "Arial";
        }
    }
}
