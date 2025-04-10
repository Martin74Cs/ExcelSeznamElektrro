using Aplikace.Sdilene;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
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
        [DllImport("oleaut32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, out object ppunk);

        public static int record = 0;

        static Exc.Application ExcelExist()
        {
            //Exc.Application excelApp = null;
            Guid clsid = new("00024500-0000-0000-C000-000000000046"); // CLSID pro Excel.Application

            int hResult = GetActiveObject(ref clsid, IntPtr.Zero, out object excelAppObj);
            if (hResult == 0)
            {
                Exc.Application excelApp = (Exc.Application)excelAppObj;
                Console.WriteLine("Excel je spuštěn.");
                return excelApp;
            }
            Console.WriteLine("Excel není spuštěn.");
            return Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;      

            //    try
            //{
            //    // Pokus o připojení k již spuštěné instanci Excelu
            //    //excelApp = (Exc.Application)Marshal.GetActiveObject("Excel.Application");
            //}
            //catch (COMException)
            //{
            //    Console.WriteLine("Excel není spuštěn.");
            //    return Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
            //}

            //if (excelApp != null)
            //{
            //    Console.WriteLine("Excel je spuštěn.");
            //    // Nyní můžete pracovat s aplikací Excel
            //    // Například první otevřený sešit
            //    Exc.Workbook workbook = excelApp.Worksheets[1];
            //    Exc.Worksheet worksheet = (Exc.Worksheet)workbook.Worksheets[1];
            //    // Proveďte nějaké operace s Excel
            //    //Console.WriteLine("Název prvního listu: " + worksheet.Name);
            //    return excelApp;
            //}
        }

        /// <summary> Vytvoření nového Excel dokumentu </summary>
        public static Exc.Workbook? VytvorNovyDokument()
        {
            //Exc.Application? App = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
            var App = ExcelExist();
            if (App == null) return null;
            App.Visible = true;

            // Vytvoření nového sešitu
            Exc.Workbook NovyDokument = App.Workbooks.Add();
            Console.Write("\nVytvořen prázný dokument Excel.");
            return NovyDokument;
        }

        public static Exc.Workbook? NovyExcelSablona(string cesta)
        {
            /// <summary> Cesta k dresaři kde bylo spuštěno nevím jak funguje u dll </summary>
            string AktuallniAdresear = System.Environment.CurrentDirectory + @"\";
            /// <summary> Cesta k Aresaři kde bylo spuštěno nevím jak funguje u dll </summary>
            string AktuallniAdresearJinak = System.IO.Directory.GetCurrentDirectory() + @"\";

            string BaseAdress = Path.Combine(System.Environment.CurrentDirectory, "Podpora");
            string sablona =  Path.Combine(BaseAdress, "Sablona_SSaZ.xlsx");
            // pokud neexistuje vlastní šablona použij výchozí
            if (!File.Exists(sablona))
               return VytvorNovyDokument();
            Exc.Application? App = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
            if (App == null) return null;
            App.Visible = true;
            if(File.Exists(cesta))
                File.Delete(cesta);
            File.Copy(sablona, cesta);
            var sesit = App.Workbooks.Open(cesta);
            Console.Write("\nVytvořen soubor ze šablony Excel.");
            return sesit;
        }


        /// <summary> Přidání nového listu do Excelového dokumentu </summary>
        public static Exc.Worksheet PridatNovyList(Exc.Workbook Dokument, string NazevListu)
        {
            Exc.Worksheet? xls = ExcelApp.GetSheet(Dokument, NazevListu);
            if(xls == null)
                // Přidání nového listu na konec sešitu
                xls = Dokument.Sheets.Add(After: Dokument.Sheets[Dokument.Sheets.Count]);
            xls.Name = NazevListu;
            xls.Activate();
            return xls;
        }

        /// <summary> nastavení listu dle jeho jména</summary>
        public static Exc.Worksheet? GetSheet(Exc.Workbook Dokument, string Nazev)
        {
            foreach (Exc.Worksheet item in Dokument.Sheets)
            {
                if (item.Name == Nazev)
                    return item;
            }
            return null;
        }

        /// <summary>Nový dokument v exelu</summary>
        public static Exc.Workbook? DokumetExcel(string Cesta)
        {
            //Exc.Application App = AplikaceExcel();
            Exc.Application? App = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
            if (App == null) return null;
            App.Visible = true;

            Console.Write("\nKontrolaOtevenehoNeboOtevreniSobroruExel - OK");
            return KontrolaOtevenehoNeboOtevreniSobroruExel(App, Cesta);
        }

        /// <summary>Kontrola otevřeného souboru v Excel</summary>
        public static Exc.Workbook KontrolaOtevenehoNeboOtevreniSobroruExel(Exc.Application ExApp, string Cesta)
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
        public static List<List<string>> ExelLoadTable(string cesta, string zalozka, int Radek, int[] CteniSloupcu, string[] TextPole)
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
                int x = 0;
                cteniPole = new List<string>();
                foreach (var item in CteniSloupcu)
                {
                    //Čtení buňky
                    Exc.Range Pok = Zal.Cells[i, item];
                    object cteni = Pok.Value;

                    string xxx = Convert.ToString(cteni);
                    if (!string.IsNullOrEmpty(xxx))
                    {
                        cteniPole.Add(xxx);
                    }
                        //object obj = new Zarizeni();
                    else
                        cteniPole.Add("");
                }
                if (!string.IsNullOrEmpty(cteniPole[1]) && cteniPole[1] != "0")
                {
                    Pole.Add(cteniPole);
                    Console.Write("\nRadek=" + i.ToString() + "\t" + cteniPole[0]);
                }

                if (i > 100 && Pole.Last().First().Count() < 2) break;
            }
            Console.Write("\nUkončení Excel");
            //Xls.Save();
            //Console.Write("\nSave OK");
            ExcelQuit(Xls);
            Console.Write("\nUkončení Excel");
            return Pole;
        }

                /// <summary> uložení dat do excel podle kriterii </summary>
        public static List<Zarizeni> ExelLoadTableTrida(string cesta, string zalozka, int Radek, int[] CteniSloupcu, string[] TextPole)
        {
            if (!System.IO.File.Exists(cesta)) return [];

            var Xls = DokumetExcel(cesta);
            if (Xls == null) return [];
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            Exc.Worksheet? Zal = GetSheet(Xls, zalozka);
            if (Zal == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.Write("\nSheet=" + Zal.Name);

            var Pole = new List<Zarizeni>();
            Console.Write("\nZal.Rows.Count=" + Zal.Rows.Count);
            var Test = new List<Zarizeni>();
            for (int i = Radek; i < Zal.Rows.Count; i++)
            {
                var obj = new Zarizeni();
                int x = 0;
                foreach (var item in CteniSloupcu)
                {
                    //Čtení buňky
                    Exc.Range Pok = Zal.Cells[i, item];
                    object cteni = Pok.Value;

                    string xxx = Convert.ToString(cteni);
                    if (!string.IsNullOrEmpty(xxx))
                    {
                        //ukladnní infomací do třídy dle jejího názvu parametru
                        Zarizeni.NastavVlastnost(obj, TextPole[x++], cteni);
                    }
                }
                Test.Add(obj);

                if (i > 100 && obj.Tag.Count() < 2) break;
            }
            Console.Write("\nUkončení Excel");
            //Xls.Save();
            //Console.Write("\nSave OK");
            ExcelQuit(Xls);
            Console.Write("..... OK");
            return Pole;
        }

        /// <summary> uložení dat do excel podle kdyterii </summary>
        public static void ExcelSaveJeden(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
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
        public static void ExcelSaveSloupec(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
        {
            string cesta1 = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
            var PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku, "Motory500V", 2, []);

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

        public static void ExcelSaveT<T>(Worksheet xls, T[] pole, string Nazev)
        {
            // Získání názvu typu T
            string className = typeof(T).Name;
            //ed.WriteMessage("\nclassName " + className);

            var TridaPole = pole.GetType();
            //ed.WriteMessage("\nTridaPole " + TridaPole.Name);

            var Sloupce = typeof(T).GetProperties();
            //ed.WriteMessage("\nSloupce " + Sloupce.Length);

            //Table tab = new Table();
            //tab.TableStyle = Sdilene.Nastav.SetTable();

            //Nastavení velikosti tabulky
            //tab.SetSize(pole.Length + 2, Sloupce.Length);
            //ed.WriteMessage("\nVelikost tabulky " + pole.Length + ", " + Sloupce.Length);

            int row = 1; int col = 1;
            xls.Cells[row, col].value = Nazev;
            row++;
            foreach (var item in Sloupce)
            {
                // Získání atributu DisplayAttribute
                DisplayAttribute displayAttribute = item.GetCustomAttributes(typeof(DisplayAttribute), false).Cast<DisplayAttribute>().FirstOrDefault();
                xls.Cells[row, col].value = item.Name.ToUpper();
                if (displayAttribute != null)
                    xls.Cells[row, col].value = displayAttribute.Name;

                //tab.Cells[row, col].TextStyleId = Sdilene.Nastav.SetROMANS();
                //tab.Columns[col].Width = (item.Name.Length * 3) + 5;
                col++;
            }
            //ed.WriteMessage("\nFunguje");
            col = 1;
            row++;
            foreach (var item in pole)
            {
                //ed.WriteMessage("\nFunguje Sloupce" + Sloupce.Length);
                foreach (var Property in Sloupce)
                {
                    //ed.WriteMessage("\nProperty.PropertyType " + Property.PropertyType);
                    //pokud je datovy typ pole
                    Console.WriteLine(Property.PropertyType.ToString());     
                    Console.WriteLine(typeof(Fluid).ToString());

                    if (Property.PropertyType == typeof(int))
                    {
                        Console.WriteLine("Jedná se o int");
                    }

                    if (Property.PropertyType.IsGenericType) 
                    {
                        Console.WriteLine("Jedná se o IsGenericType");
                        if (Property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                        {
                            var xx = item.GetType().GetProperty(Property.Name).GetValue(item) as List<T>;
                            //var dad = xx.GetProperties();

                            var Sloudvapce = typeof(T).GetProperties();
                            Console.WriteLine("Jedná se o List");
                        }
                    }


                    if (Property.PropertyType == typeof(List<>))
                    {
                        Console.WriteLine("Jedná se o IList");
                    }

                    if (Property.PropertyType == typeof(Fluid))
                    {
                        Console.WriteLine("Jedná se o seznam");
                    }

                    if (Property.PropertyType == typeof(string[]))
                    {
                        var Hodnota = (string[])item.GetType().GetProperty(Property.Name).GetValue(item);
                        string Hodpole = string.Empty;
                        //bude vytvožen seznam tohoto pole
                        foreach (var txt in Hodnota)
                        {
                            Hodpole += txt + ",";
                        }
                        xls.Cells[row, col].value = Hodpole[..^1];
                    }

                    //pokud je datovy typ string
                    if (Property.PropertyType == typeof(string))
                    {
                        //ed.WriteMessage("\nFunguje Sloupce " + Sloupce.Length);
                        //ed.WriteMessage("\nFunguje GetProperty " + Property.Name);
                        var value = item.GetType().GetProperty(Property.Name).GetValue(item).ToString(); // Získání hodnoty vlastnosti
                        //ed.WriteMessage("\nFunguje GetProperty " + value);
                        //if (value == "") 
                        //    value = "x";
                        //ed.WriteMessage("\nFunguje");
                        xls.Cells[row, col].value = value;

                        //tab.Cells[row, col].Alignment = CellAlignment.BottomLeft;
                        //tab.Cells[row, col].TextStyleId = Sdilene.Nastav.SetROMANS();
                        //tab.Columns[col].Width = (value.Length * 3) + 5;
                        col++;
                        //this.GetType().GetProperty(Property.Name).SetValue(this, Propertys);
                        //this.GetType().GetProperty(Property.Name).GetValue(Propertys);
                    }
                }
                col = 1; row++;
            }
            //tab.GenerateLayout();
            //return; //tab;
        }

        public static void NadpisMIlan(Worksheet xls)
        {
            string Nad = @"    |     |   |     |     |                                        |  |KAPACITA        |                        |        |        |      |EL.  |        ";
            int col = 1;
            int row = 1;
            foreach (var item in Nad.Split('|'))
            {
                xls.Cells[row, col++].value = item;
            }
            row++; col = 1;
            Nad = "GUID|IO/SO|NO |PS   |TAG  |NÁZEV                                   |KS|NOSTNOST        |MEDIUM                  |OBJEM   |PRŮTOK  |HMOTN.|PŘÍK.|POZNÁMKA";
            foreach (var item in Nad.Split('|'))
            {
                xls.Cells[row, col++].value = item;
            }
            //zalamování textu - pozor pokud dále řěším šírku sloupcu nesmí být zapnuto
            xls.Range[xls.Cells[1, 1], xls.Cells[2, col - 1]].WrapText = false;
            NadpisSet(xls, (row, col - 1));
        }

        public static void ExcelSave(Worksheet xls, Item[] pole)
        {
            NadpisMIlan(xls);
            //ed.WriteMessage("\nFunguje");
            int col = 1;
            int row = 3;
            Tisk(xls, pole, ref row, col);

            for (int i = 1; i < 20; i++)
                xls.Columns[i].AutoFit();

            //tab.GenerateLayout();
            return; //tab;
        }

        public static int Tisk(Worksheet xls, Item[] pole, ref int row, int col)
        {
            foreach (var item in pole)
            {
                xls.Cells[row, col++].value = item.id.ToString();
                xls.Cells[row, col++].value = item.cunit.pfx + " " +  item.cunit.num;
                xls.Cells[row, col++].value = record++.ToString();
                xls.Cells[row, col++].value = item.munit.pfx + " " + item.munit.num;
                xls.Cells[row, col++].value = item.tag;
                xls.Cells[row, col++].value = item.name;
                xls.Cells[row, col++].value = item.pcs;

                xls.Cells[row, col+4].value = item.mass;
                xls.Cells[row, col+5].value = item.power;
                xls.Cells[row, col+6].value = item.note;

                if (item.fluid.Count > 0)
                {
                    if (item.fluid.Count > 1) row++;
                    foreach (var item2 in item.fluid)
                    {
                        xls.Cells[row, col ].value = item2.parameter.value.ToString() + " " +item2.parameter.unit;
                        xls.Cells[row, col + 1].value = item2.fluid;
                        xls.Cells[row, col + 2].value = item2.volume;
                        xls.Cells[row, col + 3].value = item2.flowrate;
                        row++;
                    }
                    col += 4; row--;
                }
                else
                    col += 4;

                // Definování rozsahu pomocí čísel řádků a sloupců (např. A1:C3)
                Exc.Range range = xls.Range[xls.Cells[row, 1], xls.Cells[row, col] ];
                
                // Nastavení okrajů kolem buněk
                range.Borders[Exc.XlBordersIndex.xlEdgeBottom].LineStyle = Exc.XlLineStyle.xlContinuous;

                //Exc.Range range1 = xls.Range[xls.Cells[row, 1], xls.Cells[row, 15]];
                if (record % 2 == 1)
                    range.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

                if (item.subitem.Count > 0)
                {
                    row++; col = 1;
                    //row = Tisk(xls, item._Item__subitem.ToArray(), row, col);
                    Tisk(xls, [.. item.subitem], ref row, col);
                }
                else 
                {
                    row++; col = 1;
                }


            }
            return row;
        }


        public static void NadpisSet(Worksheet xls,  (int,int) data)
        {
            //Podtržení nadpisů
            
            // Výběr konkrétní oblasti buněk, např. A1:C3
            //Exc.Range range = ListExcel.Range["A1", "M1"];

            // Definování rozsahu pomocí čísel řádků a sloupců (např. A1:C3)
            Exc.Range range = xls.Range[
                xls.Cells[1, 1],  // A1 (1. řádek, 1. sloupec)
                xls.Cells[data.Item1, data.Item2] // Vstup (data.Item1, data.Item2)
            ];

            // Nastavení okrajů kolem buněk
            // LineStyle: Může být xlContinuous, xlDash, xlDot a další styly čar.
            range.Borders[Exc.XlBordersIndex.xlEdgeLeft].LineStyle = Exc.XlLineStyle.xlContinuous;
            range.Borders[Exc.XlBordersIndex.xlEdgeRight].LineStyle = Exc.XlLineStyle.xlContinuous;
            range.Borders[Exc.XlBordersIndex.xlEdgeTop].LineStyle = Exc.XlLineStyle.xlContinuous;
            range.Borders[Exc.XlBordersIndex.xlEdgeBottom].LineStyle = Exc.XlLineStyle.xlContinuous;

            // Další možnosti nastavení tloušťky a barvy okrajů
            //range.Borders.LineStyle = Exc.XlLineStyle.xlContinuous;

            // Weight: Určuje tloušťku čáry(xlThin, xlMedium, xlThick).
            //range.Borders.Weight = Exc.XlBorderWeight.xlMedium;  // nebo xlMedium, xlThick - tlustá

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
            //range.Orientation = 90;

            // Nastavení barvy buňky (pozadí) (např. světle modrá)
            range.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);

            //range.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(173, 216, 230));  // Světle modrá

            // Automatické přizpůsobení šířky sloupce (např. pro sloupec A)
            for (int i = 1; i <= data.Item2; i++)
                xls.Columns[i].AutoFit();

            // Automatické přizpůsobení výšky řádku (např. pro řádek 1)
            xls.Rows[1].AutoFit();
            xls.Rows[2].AutoFit();

            //range 
            //range.Columns["A:Z"].AutoFit();
            //range.Rows["1"].AutoFit();
            //range.Rows["2"].AutoFit();
        }

        public static (int,int) Nadpis(Worksheet Xls)
        {
            int col = 1;
            //Range["A1"]
            Xls.Cells[1,col++].Value = "Equipment\nnumber";

            //Range["B1"]
            Xls.Cells[1, col++].Value = "P&ID\nNumber";
            Xls.Cells[2, col -1].Value = "";

            Xls.Cells[1, col++].Value = "Equipment name";
            Xls.Cells[2, col - 1].Value = "";

            Xls.Cells[1, col++].Value = "Power(electric)\n(EU Units)";
            Xls.Cells[2, col - 1].Value = "[kW]";

            Xls.Cells[1, col++].Value = "Package unit Power";
            Xls.Cells[2, col - 1].Value = "";

            Xls.Cells[1, col++].Value = "Variable speed drive";
            Xls.Cells[2, col - 1].Value = "";

            Xls.Cells[1, col++].Value = "PROUD Z TAB. PRO 500V";
            Xls.Cells[2, col - 1].Value = "[A]";

            Xls.Cells[1, col++].Value = "Power(electric)\n(US Units)";
            Xls.Cells[2, col - 1].Value = "[HP]";

            Xls.Cells[1, col++].Value = "CURRENT FOR 480V";
            Xls.Cells[2, col - 1].Value = "[A]";

            Xls.Cells[1, col++].Value = "COPPER CABLE SIZE\n(EU Units)";
            Xls.Cells[2, col - 1].Value = "[mm2]";

            Xls.Cells[1, col++].Value = "COPPER CABLE SIZE\n(US Units)";
            Xls.Cells[2, col - 1].Value = "";

            Xls.Cells[1, col++].Value = "CABLE LENGHT";
            Xls.Cells[2, col - 1].Value = "[m]";

            Xls.Cells[1, col++].Value = "CABLE LENGHT";
            Xls.Cells[2, col - 1].Value = "[ft]";

            Xls.Cells[1, col++].Value = "DISTRIBUTOR EA/MCC";
            Xls.Cells[2, col - 1].Value = "";

            Xls.Cells[1, col++].Value = "DISTRIBUTOR NUMBER";
            Xls.Cells[2, col - 1].Value = "";

            // Povolení zalamování textu, aby nový řádek byl viditelný
            //Xls.Range["A1:M1"].WrapText = true;
            Xls.Range[Xls.Cells[1,1],Xls.Cells[2,col-1]].WrapText = true;

            return (2, col-1);
        }

        /// <summary> uložení dat do excel podle kryterii </summary>
        public static void ExcelSaveList(Worksheet Xls, List<List<string>> Vstup)
        {
            //var TextPole = new string[] { "Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
            //var PouzitProTabulku = new int[] { 3, 2, 7, 18, 1, 21, 63, 64, 61, 62, 65, 66 };

            int row = 2; int col = 1; 

            //kontrola špatného přpsaní dat souboru
            Exc.Range Kontrola = Xls.Cells[row + 1, col];
            if (!string.IsNullOrEmpty(Kontrola.Value))
            { 
                Console.WriteLine("Přepsat");
                if (Console.ReadKey().Key != ConsoleKey.A) return; 
            }


            //Čtení listu excel
            foreach (var radek in Vstup)
            {
                //Čtení radků excel
                var cteniPole = new List<string>();
                if (radek[3] != "" && radek[3] != "0")
                { 
                    row++; col=1; 
                    foreach (var item in radek)
                    {
                        //zapis qwe
                        var Zapis = Xls.Cells[row, col++];
                        if (double.TryParse(item, out double cislo))
                            Zapis.Value = cislo;
                        else 
                        {
                        //    if (item == "PU")
                        //    {
                        //        Zapis = Xls.Cells[row, col - 2];
                        //        Zapis.Value = item;
                        //    }
                        //    else
                                Zapis.Value = item;
                        }
                    }
                    Xls.Rows[row].AutoFit();
                }
            }
            Xls.Columns["A:Z"].AutoFit();
            return;
        }

        /// <summary> uložení dat do excel podle kdyterii </summary>
        public static void ExcelSaveProud(Worksheet ListExcel, List<List<string>> Vstup)
        {

            //Čtení listu excel
            for (int i = 3; i < ListExcel.UsedRange.Rows.Count; i++)
            {
                //Čtení kW
                Exc.Range Pok = ListExcel.Cells[i, 4];
                object cteni = Pok.Value;

                string xxx = Convert.ToString(cteni);
                if (double.TryParse(xxx, out double cislo))
                {
                    //Hledáni proudu z tabulky Motory500V
                    var Informace = Vstup.FirstOrDefault(x => Convert.ToDouble(x[0]) == cislo)?[1]; //.ToArray(); 
                    if (double.TryParse(Informace, out double Proud))
                    {
                        Exc.Range Zapis1 = ListExcel.Cells[i, 7];
                        Zapis1.Value = Proud;
                    }
                }

                if (cteni == null && i > 100)
                    break;
            }
            return;
        }

        /// <summary> doplnění vzorců doExel </summary>
        public static void ExcelSaveVzorce(Worksheet ListExcel, int Pocet)
        {
            //Čtení listu excel
            for (int i = 3; i < ListExcel.UsedRange.Rows.Count; i++)
            {
                // Dynamický vzorec (např. sčítání hodnot v buňkách A a B na daném řádku)
                //string formula = $"=A{row}+B{row}";
                //string formula = $"=Cells({i}, 3)+Cells({3}, 2)";
                //string formula = $"=Cells({i}, 3)*1,34102";
                //ListExcel.Cells[i, 6].Formula = formula;

                // Dynamický vzorec pomocí Excelové notace (např. C pro sloupec 3)
                //string formula = $"=C{i}*1.34102";  // C{i} odkazuje na buňku ve sloupci C (3) a řádku i

                //převod kilowatů na koně Kw -> HP * 
                ListExcel.Cells[i, 8].Formula = $"=D{i}*1.341022";

                //Převod prodů u 500 V na 480V
                ListExcel.Cells[i, 9].Formula = $"=G{i}*500/480";

                //Převod metry na stopy m -> ft 
                ListExcel.Cells[i, 13].Formula = $"=L{i}*3.280839895";

                if (i > Pocet)
                    break;
            }
            return;
        }

        /// <summary> Ze zadaného listu Exel vytvoř DataTable - podle zvolených sloupců </summary>
        public static System.Data.DataTable GetTable(Exc.Worksheet oSheet, int rowNadpis, int[] sloupec)
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
        public static bool ExcelKontrolaInstalace()
        {
            if (Type.GetTypeFromProgID("Excel.Application") != null)
                return true;
            return true;
        }

        //ukončení worksheet
        public static bool ExcelQuit(Exc.Workbook work)
        {
            // Ukončení aplikace Excel
            work.Application.Quit();

            // Uvolněte paměť
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        public static void ExcelSaveKabel(Worksheet ListExcel, List<List<string>> Vstup)
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
                    Exc.Range Zapis1 = ListExcel.Cells[i, 12];
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

        public static void ExcelSaveRozvadec(Worksheet ListExcel, List<List<string>> Vstup)
        {
            //Čtení listu excel
            //for (int i = 2; i < ListExcel.Rows.Count; i++)
            //skutečný počet použitých rádků
            for (int i = 2; i < ListExcel.UsedRange.Rows.Count; i++)
            {
                //Čtení nazvu
                Exc.Range Pok = ListExcel.Cells[i, 1];
                object cteni = Pok.Value;
                string xxx = Convert.ToString(cteni);

                //hledni proudu z tabulky delka
                var Informace = Vstup.FirstOrDefault(x => x[0] == xxx); //.ToArray(); 

                //mcc
                Exc.Range Zapis = ListExcel.Cells[i, 14];
                Zapis.Value = Informace?[8];

                //mcc
                if (double.TryParse(Informace?[9], out double cislo))
                {
                    Exc.Range Zapis1 = ListExcel.Cells[i, 15];
                    Zapis1.Value = cislo;
                }

                //if (string.IsNullOrEmpty(xxx) && i > 100)
                    //break;
            }
        }

        public static List<List<string>> ExcelLoadWorksheet(Worksheet xls, int[] pouzitProTabulku)
        {
            var Data = new List<List<string>>();
            string Cteni = "";
            //Čtení listu excel
            for (int i = 3; i < xls.UsedRange.Rows.Count; i++)
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

                if (string.IsNullOrEmpty(Cteni) && i > 100)
                    break;
            }
            return Data;
        }

        public static void ExcelSaveTable(Worksheet xls, List<List<string>> data, int Row)
        {
            Row--;
            int X1 = Row;
            int j = 1;
            foreach (var radek in data)
            {
                Console.WriteLine("Radek " + Row);
                Row++; j = 1;
                foreach (var item in radek)
                {
                    Exc.Range Zapis1 = xls.Cells[Row, j++];
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

        public static void ExcelSaveNadpis(Worksheet xls, List<List<string>> PoleData)
        {
            xls.Activate();
            Nadpis(xls, "A1:D1", "Označeni", PoleData);
            Nadpis(xls, "E1:H1", "Kabel", PoleData);
            Nadpis(xls, "I1:I1", "Zařízení", PoleData);
            Nadpis(xls, "J1:M1", "Odkud", PoleData);
            Nadpis(xls, "N1:N1", "", PoleData);
            Nadpis(xls, "O1:R1", "Kam", PoleData);
            Nadpis(xls, "S1:T1", "Delka", PoleData);

            xls.Range["G2"].Value = "[mm2]";
            xls.Range["H2"].Value = "[AWG]";
            xls.Range["S2"].Value = "[m]";
            xls.Range["T2"].Value = "[ft]";
        }

        public static void Nadpis(Worksheet xls, string pole, string Text, List<List<string>> PoleData)
        {
            // Sloučení buněk od A1 do C1
            var rada = xls.Range[pole];
            //Koontrola počtu buněk nelze sloučit jen jednu bunku.
            if (rada.Cells.Count > 1)
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

            SetFont(rada.Font);

            //Formátování nadpisů
            //Exc.Range range = xls.Range["A1", "M1"];
            // Definování rozsahu pomocí čísel řádků a sloupců (např. A1:C3)
            //Exc.Range range = xls.Range[xls.Cells[3, 1], xls.Cells[PoleData.Count(), PoleData.First().Count()]];

            string v = string.Concat(pole[..^1], (PoleData.Count() + 2).ToString());
            Exc.Range range = xls.Range[v];
            Ramecek(xls,range.Borders);
        }

        /// <summary> Nastavení Stylu písma </summary>
        public static void SetFont(Exc.Font Fonty)
        {
            // Nastavení barvy textu (např. červená)
            Fonty.Color = ColorTranslator.ToOle(Color.Red);
            //range.Font.Color = ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));  // Červená barva

            //Tučné písmo
            Fonty.Bold = true;
            //range.Font.Italic = true;

            //Velikost písma
            Fonty.Size = 14;

            //Styl pisma
            Fonty.FontStyle = "Arial";
        }

        /// <summary>Orámování rozsahu Rameček </summary>
        public static void Ramecek(Worksheet xls, Exc.Borders borders)
        {
            // Výběr rozsahu buněk (např. A1:C3)
            //Exc.Range range = xls.Range["A1", "C3"];
            //Exc.Range range = xls.Range["A1:C3"];

            // Přidání rámečku kolem vybraného rozsahu
            //Exc.Borders borders = range.Borders;

            // Nastavení stylu a tloušťky okrajů uvnitř rozsahu
            //borders.LineStyle = Exc.XlLineStyle.xlContinuous;
            //borders.Weight = Exc.XlBorderWeight.xlThin;

            // Horní hrana
            borders[Exc.XlBordersIndex.xlEdgeTop].LineStyle = Exc.XlLineStyle.xlContinuous;
            borders[Exc.XlBordersIndex.xlEdgeTop].Weight = Exc.XlBorderWeight.xlThin;

            // Spodní hrana
            borders[Exc.XlBordersIndex.xlEdgeBottom].LineStyle = Exc.XlLineStyle.xlContinuous;
            borders[Exc.XlBordersIndex.xlEdgeBottom].Weight = Exc.XlBorderWeight.xlThin;

            // Levá hrana
            borders[Exc.XlBordersIndex.xlEdgeLeft].LineStyle = Exc.XlLineStyle.xlContinuous;
            borders[Exc.XlBordersIndex.xlEdgeLeft].Weight = Exc.XlBorderWeight.xlThin;

            // Pravá hrana
            borders[Exc.XlBordersIndex.xlEdgeRight].LineStyle = Exc.XlLineStyle.xlContinuous;
            borders[Exc.XlBordersIndex.xlEdgeRight].Weight = Exc.XlBorderWeight.xlThin;

            // Pokud chcete přidat vnitřní hranice
            //borders[Exc.XlBordersIndex.xlInsideHorizontal].LineStyle = Exc.XlLineStyle.xlContinuous;
            //borders[Exc.XlBordersIndex.xlInsideHorizontal].Weight = Exc.XlBorderWeight.xlThin;

            //borders[Exc.XlBordersIndex.xlInsideVertical].LineStyle = Exc.XlLineStyle.xlContinuous;
            //borders[Exc.XlBordersIndex.xlInsideVertical].Weight = Exc.XlBorderWeight.xlThin;

        }

        /// <summary>Nový dokument Elektro pro přípravu elektro seznamů </summary>
        internal static Worksheet ExcelElektro(string cesta)
        {    
            Exc.Workbook? Doc;
            Exc.Worksheet? xls;

            if (File.Exists(cesta))
            {
                Doc = ExcelApp.DokumetExcel(cesta);
                if (Doc == null) return null;
                //Nastavení listu
                xls = ExcelApp.GetSheet(Doc, "Seznam Elektro");
                if (Doc == null) return null;
            }
            else
            { 
                Doc = VytvorNovyDokument();
                xls = PridatNovyList(Doc, "Seznam Elektro");
            }
            xls.Activate();
            return xls;
        }
    }
}
