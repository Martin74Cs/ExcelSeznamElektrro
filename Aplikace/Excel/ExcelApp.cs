using Aplikace.Sdilene;
using Aplikace.Tridy;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Excel
{
    public class ExcelApp
    {
        public Exc.Application App { get; set; } 
        public Exc.Workbook Doc { get; set; }
        public Exc.Worksheet Xls { get; set; }
        public int Process { get; set; }
        static int record = 0;

        public ExcelApp()   : this("") // zavolá druhý konstruktor s prázdným stringem
        { }

        public ExcelApp(string Cesta)
        {
            //var App = new Exc.Application
            App = new Exc.Application
            {
                Visible = true,
                DisplayAlerts = false // tohle je klíčové!
            };
            Process = Soubory.GetExcelProcess(App);

            if (File.Exists(Cesta))
            {
                Console.WriteLine("\nOpem dokument Excel.");    
                // Vytvoření nového sešitu
                //Automatikcky se vytvoří nový List1
                Doc = App.Workbooks.Add(Cesta);
                Xls = Doc.Sheets[1];
                Xls.Activate();
                return;
            }
            //Exc.Workbook Doc = App.Workbooks.Add();

            Console.WriteLine("\nVytvořen prázný dokument Excel.");
            //Automatikcky se vytvoří nový List1
            Doc = App.Workbooks.Add();
            Xls = Doc.Sheets[Doc.Sheets.Count];
            Xls.Activate();
            Console.Write("\nVztvořený dokument nastaven Aktivní.");
            //return (App, Doc);
        }

        //[DllImport("oleaut32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        //private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, out object ppunk);

        //static Exc.Application ExcelExist() 
        //{
        //    //Exc.Application excelApp = null;
        //    Guid clsid = new("00024500-0000-0000-C000-000000000046"); // CLSID pro Excel.Application
        //    try
        //    {
        //        int hResult = GetActiveObject(ref clsid, IntPtr.Zero, out object excelAppObj);
        //        if (hResult == 0)
        //        {
        //            Exc.Application excelApp = (Exc.Application)excelAppObj;
        //            Console.WriteLine("Excel je spuštěn.");
        //            return excelApp;
        //        }
        //        Console.WriteLine("Excel není spuštěn.");
        //        return Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application ?? new();      
        //    }
        //    catch (Exception)
        //    {
        //        Console.WriteLine("Chyba - Excel bude nově spuštěn.");
        //        return Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application ?? new();      
        //    }

        //    //    try
        //    //{
        //    //    // Pokus o připojení k již spuštěné instanci Excelu
        //    //    //excelApp = (Exc.Application)Marshal.GetActiveObject("Excel.Application");
        //    //}
        //    //catch (COMException)
        //    //{
        //    //    Console.WriteLine("Excel není spuštěn.");
        //    //    return Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
        //    //}

        //    //if (excelApp != null)
        //    //{
        //    //    Console.WriteLine("Excel je spuštěn.");
        //    //    // Nyní můžete pracovat s aplikací Excel
        //    //    // Například první otevřený sešit
        //    //    Exc.Workbook workbook = excelApp.Worksheets[1];
        //    //    Exc.Worksheet worksheet = (Exc.Worksheet)workbook.Worksheets[1];
        //    //    // Proveďte nějaké operace s Excel
        //    //    //Console.WriteLine("Název prvního listu: " + worksheet.Name);
        //    //    return excelApp;
        //    //}
        //}

        /// <summary> Vytvoření nového Excel dokumentu </summary>
        //public void VytvorNovyDokument()
        //{
        //    //Exc.Application? App = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) as Exc.Application;
        //    //var App = ExcelExist();
        //    //if (App == null) return null;
        //    //App.Visible = true;

        //    //var App = new Exc.Application
        //    App = new Exc.Application
        //    {
        //        Visible = true,
        //        DisplayAlerts = false // tohle je klíčové!
        //    };

        //    // Vytvoření nového sešitu
        //    //Exc.Workbook Doc = App.Workbooks.Add();
        //    Doc = App.Workbooks.Add();

        //    //Automatikcky se vytvoří nový List1
        //    Console.Write("\nVytvořen prázný dokument Excel.");
        //    //return (App, Doc);
        //}

        //public void NovyExcelSablona(string cesta)
        //{
        //    /// <summary> Cesta k dresaři kde bylo spuštěno nevím jak funguje u dll </summary>
        //    //var AktuallniAdresear = System.Environment.CurrentDirectory + @"\";
        //    /// <summary> Cesta k Aresaři kde bylo spuštěno nevím jak funguje u dll </summary>
        //    //var AktuallniAdresearJinak = System.IO.Directory.GetCurrentDirectory() + @"\";

        //    string BaseAdress = Path.Combine(System.Environment.CurrentDirectory, "Podpora");
        //    string sablona =  Path.Combine(BaseAdress, "Sablona_SSaZ.xlsx");
        //    // pokud neexistuje vlastní šablona použij výchozí
        //    //if (!File.Exists(sablona))
        //    //var (App, Doc) = VytvorNovyDokument();
        //    //VytvorNovyDokument();
        //    //if (Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) is not Exc.Application App) return null;
        //    //App.Visible = true;

        //    if (File.Exists(cesta))
        //        File.Delete(cesta);
        //    File.Copy(sablona, cesta);

        //    //new ExcelApp(cesta);
        //    //Doc = App.Workbooks.Open(cesta);
        //    //Console.Write("\nVytvořen soubor ze šablony Excel.");
        //    //return (App, Doc);
        //}


        /// <summary> Přidání nového listu do Excelového dokumentu </summary>
        public void PridatNovyList(string NazevListu)
        {
            GetSheet(NazevListu);
            //return xls;
        }

        /// <summary> nastavení listu dle jeho jména</summary>
        public void GetSheet(string Nazev)
        {
            if (Doc == null)
                return;
            foreach (Exc.Worksheet item in Doc.Sheets)
            {
                if (item.Name == Nazev)
                {
                    Xls = item;
                    Xls.Activate();
                    return;
                }
            }
            // Přidání nového listu na konec sešitu pokud je XLs praázdné
            var listy = Doc.Sheets.Count;
            Xls = Doc.Sheets.Add(After: Doc.Sheets[listy]);
            Xls.Name = Nazev;
            Xls.Activate();
            //return null;
        }

        /// <summary>Nový dokument v exelu</summary>
        public void DokumetExcel(string Cesta)
        {
            //Exc.Application App = AplikaceExcel();

            App = new Exc.Application
            {
                Visible = true,
                DisplayAlerts = false // tohle je klíčové!
            };
            Process = Soubory.GetExcelProcess(App);
            //if (Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")) is not Exc.Application App)
            //{
            //    App.Visible = true;
            //    return null;
            //} 

            Console.Write("\nKontrolaOtevenehoNeboOtevreniSobroruExel - OK");
            Doc = KontrolaOtevenehoNeboOtevreniSobroruExel(Cesta);
            //return (App, KontrolaOtevenehoNeboOtevreniSobroruExel(App, Cesta));
        }

        /// <summary>Kontrola otevřeného souboru v Excel</summary>
        public Exc.Workbook KontrolaOtevenehoNeboOtevreniSobroruExel(string Cesta)
        {
            Console.Write("\nMetoda Kontrola Oteveneho Nebo Otevreni Sobroru Exel");
            Console.Write("\nCesta" + Cesta.ToLowerInvariant());
            if (File.Exists(Cesta))
            {
                Console.Write("\nSoubor není otevřen kontrola ");
                //return null;
                //nefunuguje otevření souboru
                return App.Workbooks.Open(Cesta.ToLowerInvariant());
            }
            return App.Workbooks.Add();
            //foreach (Exc.Workbook item in App.Workbooks)
            //{
            //    Console.Write("\nName=" + item.Name);
            //    if (item.Name == System.IO.Path.GetFileName(Cesta.ToLowerInvariant()))
            //        return item;
            //}

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

            //var (App,Xls) = DokumetExcel(cesta);
            DokumetExcel(cesta);
            if (Xls == null) return [];
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            GetSheet(zalozka);
            if (Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.Write("\nSheet=" + Xls.Name);

            //var cteniPole = new List<string>() ;
            var Pole = new List<List<string>>();
            Console.Write("\nZal.Rows.Count=" + Xls.UsedRange.Rows.Count);
            for (int i = Radek; i < Xls.UsedRange.Rows.Count; i++)
            {
                //int x = 0;
                var cteniPole = new List<string>();
                //čtení jednotlivých řádků excelu
                foreach (var item in CteniSloupcu)
                {
                    //Čtení buňky
                    Exc.Range Pok = Xls.Cells[i, item]; 
                    object cteni = Pok.Value;

                    string xxx = Convert.ToString(cteni) ?? string.Empty;
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
                //Pojistka
                if (i > 100 && Pole.Last().First().Length < 2) break;
            }
            Console.Write("\nUkončení Excel");
            //Xls.Save();
            //Console.Write("\nSave OK");
            ExcelQuit();
            Console.Write("\nUkončení Excel");
            return Pole;
        }

        /// <summary> uložení dat do excel podle kriterii </summary>
        public List<Zarizeni> ExelTable(int Radek, string Tabulka)
        {
            //Nastavení listu
            GetSheet(Tabulka);

            //Sloupce Exel odpovídající názvům tříd.
            var dir = new Dictionary<int, string>() {
                {1, "Radek"     },
                {2, "Tag"       },
                {4, "Popis"     },
                {11, "Menic"    },
                {10, "Prikon"   },
                {18, "BalenaJednotka"   },
            };
            int pocet = 1;
            //string prikon = string.Empty;
            //var cteniPole = new List<string>();
            //var Pole = new List<List<string>>();
            var Pole = new List<Zarizeni>();
            Console.WriteLine($"[Rows.Col]=[{Xls.UsedRange.Rows.Count},{Xls.UsedRange.Columns.Count}]");
            for ( int i = Radek; i < Xls.UsedRange.Rows.Count; i++)
            {
                //int x = 0;
                //čteniPole = [];
                //čtení jednotlivých řádků excelu
                var jeden = new Zarizeni();
                for (int j = 1; j < Xls.UsedRange.Columns.Count; j++)
                {
                    //přeskok příkon prazdný
                    var prikon = Convert.ToString(Xls.Cells[i, 10].Value);
                    if (string.IsNullOrEmpty(prikon) || prikon == "0") { jeden.Prikon = ""; break; } 

                    //Čtení buňky
                    Exc.Range Pok = Xls.Cells[i, j];
                    string xxx = Convert.ToString(Pok.Value);

                    //přeskočit prázdné buňky a nulové
                    if (string.IsNullOrEmpty(xxx) || xxx == "0")
                    { 
                        //Console.WriteLine($"Radek {pocet++} - je prázdný");
                        continue;
                    }

                    //ukladnní infomací do třídy dle jejího názvu parametru
                    if (dir.TryGetValue(j, out var value))
                        jeden[value] = xxx; 
                }
                if (!string.IsNullOrEmpty(jeden.Prikon))
                {
                    jeden.Apid = ExcelLoad.Apid();
                    Pole.Add(jeden);
                    Console.WriteLine($"Radek {pocet++} - přídán");
                }
                else
                {
                    Console.WriteLine($"Radek {pocet++} - přeskočen, Příkon {jeden.Prikon} - není číslo");
                }

                //if (!string.IsNullOrEmpty(Pole[1]) && Pole[1] != "0")
                //{
                //    Pole.Add(Pole);
                //    Console.Write("\nRadek=" + i.ToString() + "\t" + Pole[0]);
                //}
                //Pojistka
                //if (i > 100 && Pole.Last().First().Length < 2) break;
            }
            Console.WriteLine("Zavřít sešit Excel");
            //Xls.Save();
            //Zal.Parent.Close();
            //Console.Write("\nSave OK");
            //ExcelQuit(Xls);
            //Console.Write("\nUkončení Excel");
            return Pole;
        }

        /// <summary> uložení dat do excel podle kriterii </summary>
        public void ClassToExcel<T>(int Row, List<T> Pole, Dictionary<int, string> Sloupce)
        {

            //var properties = typeof(T).GetProperties();
            var properties = typeof(T).GetProperties().ToDictionary(p => p.Name);

            // kontrola vlastností v dir, jestli existují v T
            foreach (var kvp in Sloupce)
            {
                if (!properties.ContainsKey(kvp.Value))
                {
                    Console.WriteLine($"[WARN] Vlastnost '{kvp.Value}' (pro sloupec {kvp.Key}) neexistuje v typu {typeof(T).Name}");
                }
            }

            // Vyfiltruj jen ty položky, které mají odpovídající vlastnost ve třídě T
            var dirFiltered = Sloupce
                .Where(kvp => properties.ContainsKey(kvp.Value))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            //var cteniPole = new List<string>();
            //var Pole = new List<List<string>>();
            //int Row = 3; 
            Console.WriteLine($"[Rows.Col]=[{Xls.UsedRange.Rows.Count},{Xls.UsedRange.Columns.Count}]");
            foreach (var item in Pole)
            {
                foreach (var kvp in dirFiltered)
                {
                    int Col = kvp.Key;
                    string propName = kvp.Value;

                    Exc.Range Pok = Xls.Cells[Row, Col];
                    var prop = properties[propName];
                    Pok.Value = prop.GetValue(item);
                }
                Row++;
            }
                for (int i = 1; i < 15; i++)
                    Xls.Columns[i].AutoFit();
            return;
        }


        /// <summary> uložení dat do excel podle kriterii </summary>
        public List<Zarizeni> ExelLoadTableTrida(string cesta, string zalozka, int Radek, int[] CteniSloupcu, string[] TextPole)
        {
            if (!System.IO.File.Exists(cesta)) return [];

            //var (App, Xls) = DokumetExcel(cesta);
            DokumetExcel(cesta);
            if (Xls == null) return [];
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            GetSheet(zalozka);
            if (Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.Write("\nSheet=" + Xls.Name);

            var Pole = new List<Zarizeni>();
            Console.Write("\nZal.Rows.Count=" + Xls.Rows.Count);
            var Test = new List<Zarizeni>();
            for (int i = Radek; i < Xls.Rows.Count; i++)
            {
                var obj = new Zarizeni();
                int x = 0;
                foreach (var item in CteniSloupcu)
                {
                    //Čtení buňky
                    Exc.Range Pok = Xls.Cells[i, item];
                    object cteni = Pok.Value;

                    string xxx = Convert.ToString(cteni)?? string.Empty;
                    if (!string.IsNullOrEmpty(xxx))
                    {
                        //ukladnní infomací do třídy dle jejího názvu parametru
                        //Zarizeni.NastavVlastnost(obj, TextPole[x++], cteni);
                        obj[TextPole[x++]] = cteni;
                    }
                }
                Test.Add(obj);

                //if (i > 100 && obj.Tag.Length < 2) break;
            }
            Console.Write("\nUkončení Excel");
            //Xls.Save();
            //Console.Write("\nSave OK");
            ExcelQuit();
            Console.Write("..... OK");
            return Pole;
        }

        /// <summary> uložení dat do excel podle kdyterii </summary>
        public void ExcelSaveJeden(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
        {
            if (!System.IO.File.Exists(cesta)) return;

            //var (App, Xls) = DokumetExcel(cesta);
            DokumetExcel(cesta);
            if (Xls == null) return;
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            GetSheet(zalozka);
            if (Xls == null) { Console.Write("\nChyba KONEC");  return; }
            //Exc.Worksheet Zal = Xls.Worksheets[zalozka];
            Console.Write("\nSheet=" + Xls.Name);

            //Čtení listu excel
            for (int i = 7; i < Xls.Rows.Count; i++)
            {
                //Čtení radků excel
                var cteniPole = new List<string>();
                foreach (var item in SloupceCteni)
                {
                    //Čtení buňky Tag
                    Exc.Range Pok = Xls.Cells[i, item];
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
                    Exc.Range Zapis = Xls.Cells[i, SloupceZapisu.First()];
                    Zapis.Value = Shoda.First();

                    //Posledni
                    Zapis = Xls.Cells[i, SloupceZapisu.Last()];
                    Zapis.Value = Shoda[8] + " " + Shoda.Last();
                }
                else
                {
                    //nebyla shoda
                    //zapis buňky
                    Exc.Range Zapis = Xls.Cells[i, SloupceZapisu.First()];
                    Zapis.Value = "Nenalezeno";
                }

                { Console.Write("\nShoda buňky " + i); }
                if (i > 500) break;
            }
            Console.Write("\nUkončení Excel");
            Doc.Save();

            Console.Write("\nSave OK");
            //Xls.Close();
            //ed.WriteMessage("\nClose OK");
            ExcelQuit();
            Console.Write("\nUkončení Excel");
            return;
        }

        /// <summary> uložení dat do excel podle kdyterii </summary>
        public void ExcelSaveSloupec(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
        {
            string cesta1 = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
            var PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku, "Motory500V", 2);
            Motory500.Vypis();

            if (!System.IO.File.Exists(cesta)) return;

            //var (App, Xls) = DokumetExcel(cesta);
            DokumetExcel(cesta);
            if (Xls == null) return;
            Console.Write("\nDokument excel - Otevřen");

            //Nastavení listu
            GetSheet(zalozka);
            if (Xls == null) return;
            //Exc.Worksheet Zal = Xls.Worksheets[zalozka];
            Console.Write("\nSheet=" + Xls.Name);

            //Čtení listu excel
            for (int i = 7; i < Xls.Rows.Count; i++)
            {
                //Čtení radků excel
                var cteniPole = new List<string>();
                foreach (var item in SloupceCteni)
                {
                    //Čtení buňky
                    Exc.Range Pok = Xls.Cells[i, item];
                    object cteni = Pok.Value;

                    //string xxx = Convert.ToString(cteni);
                    string xxx = Convert.ToString(cteni) ?? string.Empty;;
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
                            Exc.Range Zapis1 = Xls.Cells[i, SloupceZapisu.First()];
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
                    var Zapis = Xls.Cells[i, SloupceZapisu[1]];
                    Zapis.Value = Shoda[8];

                    //Rozvaděč
                    Zapis = Xls.Cells[i, SloupceZapisu[2]];
                    Zapis.Value = Shoda[9];

                    //zapis delka
                    Zapis = Xls.Cells[i, SloupceZapisu[3]];
                    if (double.TryParse(Shoda[4], out double delka))
                        Zapis.Value = delka;
                    else
                        Zapis.Value = Shoda[4].ToString();

                    //zapis AWG
                    Zapis = Xls.Cells[i, SloupceZapisu[5]];
                    if (double.TryParse(Shoda[5], out double AWG))
                        Zapis.Value = AWG;
                    else
                        Zapis.Value = Shoda[5].ToString();

                    //zapis mm2
                    Zapis = Xls.Cells[i, SloupceZapisu[4]];
                    if (double.TryParse(Shoda[10], out double mm2))
                        Zapis.Value = mm2;
                    else
                        Zapis.Value = "";
                }
                else
                {
                    //nebyla shoda
                    //zapis buňky
                    Exc.Range Zapis = Xls.Cells[i, SloupceZapisu.First()];
                    Zapis.Value = "Nenalezeno";
                }

                { Console.Write("\nShoda buňky " + i); }
                if (i > 500) break;
            }
            Console.Write("\nUkončení Excel");
            Doc.Save();

            Console.Write("\nSave OK");
            //Xls.Close();
            //ed.WriteMessage("\nClose OK");
            ExcelQuit();
            Console.Write("\nUkončení Excel");
            return;
        }

        public void ExcelSaveT<T>(T[] pole, string Nazev)
        {
            // Získání názvu typu T
            string ClassName = typeof(T).Name;
            Console.WriteLine(ClassName);

            // Získání názvu Type
            var TridaPole = pole.GetType();
            Console.WriteLine(TridaPole.Name);

            //Sada vlasnotnotí 
            var Sloupce = typeof(T).GetProperties();
            foreach (var item in Sloupce)
                Console.WriteLine(item.Name, item?.GetValue(item)?.ToString());
            //Table tab = new Table();
            //tab.TableStyle = Sdilene.Nastav.SetTable();

            //Nastavení velikosti tabulky
            //tab.SetSize(pole.Length + 2, Sloupce.Length);
            //ed.WriteMessage("\nVelikost tabulky " + pole.Length + ", " + Sloupce.Length);

            int row = 1; int col = 1;
            Xls.Cells[row, col].value = Nazev;
            row++;
            foreach (var item in Sloupce)
            {
                // Získání atributu DisplayAttribute
                DisplayAttribute displayAttribute = item.GetCustomAttributes(typeof(DisplayAttribute), false).Cast<DisplayAttribute>().FirstOrDefault();
                Xls.Cells[row, col].value = item.Name.ToUpper();
                if (displayAttribute != null)
                    Xls.Cells[row, col].value = displayAttribute.Name;

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
                    Console.WriteLine(typeof(Fluids).ToString());

                    if (Property.PropertyType == typeof(int))
                    {
                        Console.WriteLine("Jedná se o int");
                    }

                    if (Property.PropertyType.IsGenericType) 
                    {
                        Console.WriteLine("Jedná se o IsGenericType");
                        if (Property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                        {
                            var xx = item?.GetType().GetProperty(Property.Name)?.GetValue(item) as List<T>;
                            //var dad = xx.GetProperties();

                            var Sloudvapce = typeof(T).GetProperties();
                            Console.WriteLine("Jedná se o List");
                        }
                    }


                    if (Property.PropertyType == typeof(List<>))
                    {
                        Console.WriteLine("Jedná se o IList");
                    }

                    if (Property.PropertyType == typeof(Fluids))
                    {
                        Console.WriteLine("Jedná se o seznam");
                    }

                    if (Property.PropertyType == typeof(string[]))
                    {
                        //var prop = item?.GetType().GetProperty(Property.Name);
                        //var value = prop?.GetValue(item) as string[];
                        //var Hodnota = value ?? [];

                        var Hodnota = (string[])item?.GetType().GetProperty(Property.Name)?.GetValue(item);
                        string Hodpole = string.Empty;
                        //bude vytvožen seznam tohoto pole
                        foreach (var txt in Hodnota.ToString() ?? string.Empty)
                        {
                            Hodpole += txt + ",";
                        }
                        Xls.Cells[row, col].value = Hodpole[..^1];
                    }

                    //pokud je datovy typ string
                    if (Property.PropertyType == typeof(string))
                    {
                        //ed.WriteMessage("\nFunguje Sloupce " + Sloupce.Length);
                        //ed.WriteMessage("\nFunguje GetProperty " + Property.Name);
                        var value = item?.GetType().GetProperty(Property.Name)?.GetValue(item)?.ToString(); // Získání hodnoty vlastnosti
                        //ed.WriteMessage("\nFunguje GetProperty " + value);
                        //if (value == "") 
                        //    value = "x";
                        //ed.WriteMessage("\nFunguje");
                        Xls.Cells[row, col].value = value;

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

        public void NadpisMIlan()
        {
            string Nad = @"    |     |   |     |     |                                        |  |KAPACITA        |                        |        |        |      |EL.  |        ";
            int col = 1;
            int row = 1;
            foreach (var item in Nad.Split('|'))
            {
                Xls.Cells[row, col++].value = item;
            }
            row++; col = 1;
            Nad = "GUID|IO/SO|NO |PS   |TAG  |NÁZEV                                   |KS|NOSTNOST        |MEDIUM                  |OBJEM   |PRŮTOK  |HMOTN.|PŘÍK.|POZNÁMKA";
            foreach (var item in Nad.Split('|'))
            {
                Xls.Cells[row, col++].value = item;
            }
            //zalamování textu - pozor pokud dále řěším šírku sloupcu nesmí být zapnuto
            var range = Xls.Range[Xls.Cells[1, 1], Xls.Cells[2, col - 1]];
            range.WrapText = false;
            NadpisSet(range);
        }

        public void ExcelSave(Item[] pole)
        {
            NadpisMIlan();
            //ed.WriteMessage("\nFunguje");
            int col = 1;
            int row = 3;
            Tisk(pole, ref row, col);

            for (int i = 1; i < 20; i++)
                Xls.Columns[i].AutoFit();

            //tab.GenerateLayout();
            return; //tab;
        }

        public int Tisk(Item[] pole, ref int row, int col)
        {
            foreach (var item in pole)
            {
                Xls.Cells[row, col++].value = item.Id.ToString();
                Xls.Cells[row, col++].value = item.Cunit.Pfx + " " +  item.Cunit.Num;
                Xls.Cells[row, col++].value = record++.ToString();
                Xls.Cells[row, col++].value = item.Munit.Pfx + " " + item.Munit.Num;
                Xls.Cells[row, col++].value = item.Tag;
                Xls.Cells[row, col++].value = item.Name;
                Xls.Cells[row, col++].value = item.Pcs;

                Xls.Cells[row, col+4].value = item.Mass;
                Xls.Cells[row, col+5].value = item.Power;
                Xls.Cells[row, col+6].value = item.Note;

                if (item.Fluid.Count > 0)
                {
                    if (item.Fluid.Count > 1) row++;
                    foreach (var item2 in item.Fluid)
                    {
                        Xls.Cells[row, col ].value = item2.Parameter.Value.ToString() + " " +item2.Parameter.Unit;
                        Xls.Cells[row, col + 1].value = item2.Fluid;
                        Xls.Cells[row, col + 2].value = item2.Volume;
                        Xls.Cells[row, col + 3].value = item2.Flowrate;
                        row++;
                    }
                    col += 4; row--;
                }
                else
                    col += 4;

                // Definování rozsahu pomocí čísel řádků a sloupců (např. A1:C3)
                Exc.Range range = Xls.Range[Xls.Cells[row, 1], Xls.Cells[row, col] ];
                
                // Nastavení okrajů kolem buněk
                range.Borders[Exc.XlBordersIndex.xlEdgeBottom].LineStyle = Exc.XlLineStyle.xlContinuous;

                //Exc.Range range1 = xls.Range[xls.Cells[row, 1], xls.Cells[row, 15]];
                if (record % 2 == 1)
                    range.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

                if (item.Subitem.Count > 0)
                {
                    row++; col = 1;
                    //row = Tisk(xls, item._Item__subitem.ToArray(), row, col);
                    Tisk([.. item.Subitem], ref row, col);
                }
                else 
                {
                    row++; col = 1;
                }


            }
            return row;
        }

        public static void NadpisSet(Exc.Range range)
        {
            //Podtržení nadpisů
            
            // Výběr konkrétní oblasti buněk, např. A1:C3
            //Exc.Range range = ListExcel.Range["A1", "M1"];

            // Definování rozsahu pomocí čísel řádků a sloupců (např. A1:C3)
            //Exc.Range range = xls.Range[
            //    xls.Cells[1, 1],  // A1 (1. řádek, 1. sloupec)
            //    xls.Cells[data.Item1, data.Item2] // Vstup (data.Item1, data.Item2)
            //];
            

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

            SetFontRed(range.Font);

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
            //for (int i = 1; i <= data.Item2; i++)
            //    xls.Columns[i].AutoFit();

            //Nastavení sloupců hromadně
            //range.EntireColumn.AutoFit();
            foreach (Exc.Range item in range.Columns)
                item.AutoFit();

            // Automatické přizpůsobení výšky řádku (např. pro řádek 1)
            foreach (Exc.Range item in range.Rows)
                item.AutoFit();

            //xls.Rows[1].AutoFit();
            //xls.Rows[2].AutoFit();

            //range 
            //range.Columns["A:Z"].AutoFit();
            //range.Rows["1"].AutoFit();
            //range.Rows["2"].AutoFit();
        }

        public Exc.Range Nadpisy(Nadpis[] data)
        {
            
            int col = 1;
            //Tisk pole data
            foreach (var item in data)
            {
                Xls.Cells[1, col].Value = item.Name;
                Xls.Cells[2, col++].Value = item.Jednotky;
            }

            // Povolení zalamování textu, aby nový řádek byl viditelný
            //Xls.Range["A1:M1"].WrapText = true;
            var Range = Xls.Range[Xls.Cells[1, 1], Xls.Cells[2, col - 1]];
            //polovlit zalamování
            Range.WrapText = false;
            return Range;
        }



        public Exc.Range Nadpisy<T>()
        {
            int col = 1;
            //var props = typeof(T).GetProperties();

            //var properties = typeof(T).GetProperties();
            var properties = typeof(T).GetProperties().ToDictionary(p => p.Name);

            // kontrola vlastností v dir, jestli existují v T
            foreach (var kvp in Mistnost.Sloupce)
            {
                if (!properties.ContainsKey(kvp.Value))
                {
                    Console.WriteLine($"[WARN] Vlastnost '{kvp.Value}' (pro sloupec {kvp.Key}) neexistuje v typu {typeof(T).Name}");
                }
            }

            // Vyfiltruj jen ty položky, které mají odpovídající vlastnost ve třídě T
            var dirFiltered = Mistnost.Sloupce
                .Where(kvp => properties.ContainsKey(kvp.Value))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            //Tisk pole data
            foreach (var kvp in dirFiltered)
            {
                //převod na PropertyInfo
                var prop = properties[kvp.Value];

                // Načti atribut [Display(Name = "...")]
                var displayAttr = prop.GetCustomAttribute<DisplayAttribute>();
                string displayName = displayAttr?.Name ?? prop.Name;

                // Načti atribut [Jednotky("...")] - volitelně, pokud máš
                var jednotkyAttr = prop.GetCustomAttribute<JednotkyAttribute>();
                string jednotky = jednotkyAttr?.Text ?? "";

                Xls.Cells[1, col].Value = displayName;
                Xls.Cells[2, col].Value = jednotky;
                col++;
            }

            // Povolení zalamování textu, aby nový řádek byl viditelný
            //Xls.Range["A1:M1"].WrapText = true;
            var Range = Xls.Range[Xls.Cells[1, 1], Xls.Cells[2, col - 1]];
            //polovlit zalamování
            Range.WrapText = false;
            return Range;
        }

        /// <summary> uložení dat do excel podle kryterii </summary>
        public void ExcelSaveList(List<List<string>> Vstup)
        {
            //var TextPole = new string[] { "Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
            //var PouzitProTabulku = new int[] { 3, 2, 7, 18, 1, 21, 63, 64, 61, 62, 65, 66 };

            int row = 2; int col = 1; 

            //kontrola špatného přepsaní dat souboru
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
                //if (radek[3] != "" && radek[3] != "0")
                //{ 
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
                  //  }
                    Xls.Rows[row].AutoFit();
                }
            }
            //zalomení
            Xls.Columns["A:Z"].AutoFit();
            return;
        }
        /// <summary> uložení dat do excel podle kryterii </summary>
        public void ExcelSaveClass(List<Zarizeni> Vstup)
        {
            //var TextPole = new string[] { "Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC" };
            //var PouzitProTabulku = new int[] { 3, 2, 7, 18, 1, 21, 63, 64, 61, 62, 65, 66 };
            int row = 2; int col;

            //Čtení listu tridy
            foreach (var radek in Vstup)
            {
                //Čtení radků excel
                //var cteniPole = new List<string>();
                //if (radek[3] != "" && radek[3] != "0")
                //{ 
                              
                row++; col = 1;
                //zapis qwe
                var Zapis = Xls.Cells[row, col++];
                switch (col)
                {
                    case 1: 
                        Zapis.value = radek.Tag;    
                        break;
                    case 2:
                        //pid
                        Zapis.value = radek.PID;
                        break;
                    case 3:
                        //popis
                        Zapis.value = radek.Popis;
                        break;
                    case 4:
                        //prikon
                        Zapis.value = radek.Prikon;
                        break;
                    case 5:
                        //balená jednotka
                        Zapis.value = radek.BalenaJednotka;
                        break;
                    case 6:
                        //menic
                        Zapis.value = radek.Menic;
                        break;
                    case 7:
                        //proud
                        Zapis.value = radek.Proud;
                        break;
                    case 8:
                        //HP
                        Zapis.value = radek.HP;
                        break;
                    case 9:
                        //proud
                        if (double.TryParse(radek.Proud, out double proud1))
                            Zapis.value = (proud1 * 500 / 480).ToString();
                        break;
                    case 10:
                        //PruzezMM2
                        Zapis.value = radek.PruzezMM2;
                        break;
                    case 11:
                        //Pruzez US unit
                        //Zapis.value = radek.PruzezMM2;
                        break;
                    case 12:
                        //delka
                        Zapis.value = radek.Delka;
                        break;
                    case 13:
                        //delka stopy
                        //Zapis.value = radek.Delka;
                        break;
                    case 14:
                        //royvaděč
                        Zapis.value = radek.Rozvadec;
                        break;
                    case 15:
                        //royvaděč
                        Zapis.value = radek.RozvadecCislo;
                        break;
                    default:
                        break;
                }

                //if (double.TryParse(item, out double cislo))
                //    Zapis.Value = cislo;
                //else
                //{
                //    Zapis.Value = item;
                //}
                //  }
                Xls.Rows[row].AutoFit();
            }
            //zalomení
            Xls.Columns["A:Z"].AutoFit();
            return;
        }

        /// <summary> uložení dat do excel podle kryterii </summary>
        public void ExcelSaveProud(List<List<string>> Vstup)
        {

            //Čtení listu excel
            for (int i = 3; i < Xls.UsedRange.Rows.Count; i++)
            {
                //Čtení kW
                Exc.Range Pok = Xls.Cells[i, 4];
                object cteni = Pok.Value;

                string xxx = Convert.ToString(cteni) ?? string.Empty;
                if (double.TryParse(xxx, out double cislo))
                {
                    //Hledáni proudu z tabulky Motory500V
                    var Informace = Vstup.FirstOrDefault(x => Convert.ToDouble(x[0]) == cislo)?[1]; //.ToArray(); 
                    if (double.TryParse(Informace, out double Proud))
                    {
                        Exc.Range Zapis1 = Xls.Cells[i, 7];
                        Zapis1.Value = Proud;
                    }
                }

                if (cteni == null && i > 100)
                    break;
            }
            return;
        }

        /// <summary> doplnění vzorců doExel </summary>
        public void ExcelSaveVzorce(int Pocet)
        {
            //Čtení listu excel
            for (int i = 3; i < Xls.UsedRange.Rows.Count; i++)
            {
                // Dynamický vzorec (např. sčítání hodnot v buňkách A a B na daném řádku)
                //string formula = $"=A{row}+B{row}";
                //string formula = $"=Cells({i}, 3)+Cells({3}, 2)";
                //string formula = $"=Cells({i}, 3)*1,34102";
                //ListExcel.Cells[i, 6].Formula = formula;

                // Dynamický vzorec pomocí Excelové notace (např. C pro sloupec 3)
                //string formula = $"=C{i}*1.34102";  // C{i} odkazuje na buňku ve sloupci C (3) a řádku i

                //převod kilowatů na koně Kw -> HP * 
                Xls.Cells[i, 8].Formula = $"=D{i}*1.341022";

                //Převod prodů u 500 V na 480V
                Xls.Cells[i, 9].Formula = $"=G{i}*500/480";

                //Převod metry na stopy m -> ft 
                Xls.Cells[i, 13].Formula = $"=L{i}*3.280839895";

                if (i > Pocet)
                    break;
            }
            return;
        }

        /// <summary> Ze zadaného listu Exel vytvoř DataTable - podle zvolených sloupců </summary>
        public System.Data.DataTable GetTable(int rowNadpis, int[] sloupec)
        {
            //ed.WriteMessage("\nZačala metoda GetTable");
            //ed.WriteMessage("\nNadpis=" + rowNadpis + ", sloupec=" + sloupec.Length + ", Name=" + oSheet.Name + ", Rows=" + oSheet.Rows.Count);

            var Table = new System.Data.DataTable("Tabulka");

            // Načtěte konkrétní řádek
            Exc.Range rowRange = Xls.Rows[rowNadpis];
            //ed.WriteMessage("\nVelikost Sheet.Rows " + rowRange.Columns.Count); //vysledek je 16384

            Exc.Range range = Xls.UsedRange;
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

                //DataRow range;
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

        
        ///<summary>Console.WriteLine("Zavřit dokument ");ukončení worksheet </summary>
        public bool ExcelQuit(bool UkonceniApplikace = true)
        {
            //ukončení worksheet
            if (Doc == null) return false;
            Console.WriteLine("Zavřit dokument "); 
            //Uložení dokumentu
            Doc.Close();

            if (UkonceniApplikace == true)
            { 
                // Ukončení aplikace Excel
                Console.WriteLine("Zavřit dokument "); 
                if (App == null) return false;
                App.Quit();
            }

            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                Marshal.ReleaseComObject(Doc);
                Marshal.ReleaseComObject(App);

                // Uvolněte paměť
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            Soubory.KillExcel(Process);

            return true;
        }

        public void ExcelSaveKabel(List<List<string>> Vstup)
        {

            //Čtení listu excel
            for (int i = 2; i < Xls.Rows.Count; i++)
            {
                //Čtení nazvu
                Exc.Range Pok = Xls.Cells[i, 1];
                object cteni = Pok.Value;

                string xxx = Convert.ToString(cteni) ?? string.Empty;

                //hledni proudu z tabulky delka
                var Informace = Vstup.FirstOrDefault(x => x[0] == xxx); //.ToArray(); 

                //delka
                if (double.TryParse(Informace?[4], out double delka))
                {
                    Exc.Range Zapis1 = Xls.Cells[i, 12];
                    Zapis1.Value = delka;
                }

                //awg
                Exc.Range Zapis = Xls.Cells[i, 11];
                Zapis.Value = Informace?[5];

                //mm2
                if (double.TryParse(Informace?[10], out double mm2))
                {
                    Exc.Range Zapis1 = Xls.Cells[i, 10];
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
                string xxx = Convert.ToString(cteni) ?? string.Empty;

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

        public List<List<string>> ExcelLoadWorksheet(int[] pouzitProTabulku)
        {
            var Data = new List<List<string>>();
            string Cteni = "";
            //Čtení listu excel
            for (int i = 3; i < Xls.UsedRange.Rows.Count; i++)
            {
                var Radek = new List<string>();
                foreach (var item in pouzitProTabulku)
                {
                    //zapis qwe
                    var Zapis = Xls.Cells[i, item];
                    Cteni = Convert.ToString(Zapis.Value);
                    Radek.Add(Cteni);
                }
                Data.Add(Radek);

                if (string.IsNullOrEmpty(Cteni) && i > 100)
                    break;
            }
            return Data;
        }

        public void ExcelSaveTable(List<List<string>> data, int Row)
        {
            Row--;
            int j;
            foreach (var radek in data)
            {
                Console.WriteLine("Radek " + Row);
                Row++; j = 1;
                foreach (var item in radek)
                {
                    Exc.Range Zapis1 = Xls.Cells[Row, j++];
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

        public void ExcelSaveNadpis(List<List<string>> Ramecek)
        {
            Xls.Activate();
            Nadpis("A1:D1", "Označeni", Ramecek);
            Nadpis("E1:H1", "Kabel", Ramecek);
            Nadpis("I1:I1", "Zařízení", Ramecek);
            Nadpis("J1:M1", "Odkud", Ramecek);
            Nadpis("N1:N1", "", Ramecek);
            Nadpis("O1:R1", "Kam", Ramecek);
            Nadpis("S1:S1", "Delka", Ramecek);

            //xls.Range["G2"].Value = "[mm2]";
            //xls.Range["H2"].Value = "[AWG]";
            //xls.Range["S2"].Value = "[m]";
            //xls.Range["T2"].Value = "[ft]";
        }

        public void Nadpis(string pole, string Text)
        {
            Nadpis(pole, Text, []);
        }

        public void Nadpis(string pole, string Text, List<List<string>> PoleData)
        {
            // Sloučení buněk od A1 do C1
            var range = Xls.Range[pole];
            //Koontrola počtu buněk nelze sloučit jen jednu bunku.
            if (range.Cells.Count > 1)
                range.Merge();

            // Nastavení auto šířky sloupce
            range.WrapText = false;

            //Hodnota bunky
            range.Value = Text;

            //zarovnání
            range.HorizontalAlignment = Exc.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Exc.XlVAlign.xlVAlignCenter;

            // Další možnosti nastavení tloušťky a barvy okrajů
            range.Borders.LineStyle = Exc.XlLineStyle.xlContinuous;

            // Weight: Určuje tloušťku čáry(xlThin, xlMedium, xlThick).
            range.Borders.Weight = Exc.XlBorderWeight.xlMedium;  // nebo xlMedium, xlThick - tlusté

            //Color: Převádí barvu z knihovny System.Drawing.Color na formát použitelný v Excelu.
            range.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black); // nastavení barvy čar

            SetFont(range.Font);

            //Formátování nadpisů
            //Exc.Range range = xls.Range["A1", "M1"];
            // Definování rozsahu pomocí čísel řádků a sloupců (např. A1:C3)
            //Exc.Range range = xls.Range[xls.Cells[3, 1], xls.Cells[PoleData.Count(), PoleData.First().Count()]];

            string v = string.Concat(pole[..^1], (PoleData.Count + 2).ToString());
            range = Xls.Range[v];
            Ramecek(range.Borders);
        }

        /// <summary> Nastavení Stylu písma </summary>
        public static void SetFont(Exc.Font Fonty)
        {
            // Nastavení barvy textu (např. červená)
            //Fonty.Color = ColorTranslator.ToOle(Color.Red);
            //range.Font.Color = ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));  // Červená barva

            //Tučné písmo
            Fonty.Bold = true;
            //range.Font.Italic = true;

            //Velikost písma
            Fonty.Size = 14;

            //Styl pisma
            Fonty.FontStyle = "Arial";
        }

        /// <summary> Nastavení Stylu písma červené</summary>
        public static void SetFontRed(Exc.Font Fonty)
        {
            Fonty.Color = ColorTranslator.ToOle(Color.Red);
            SetFont(Fonty);
        }
        /// <summary>Orámování rozsahu Rameček </summary>
        public static void Ramecek(Exc.Borders borders)
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
        //public void ExcelElektro(string cesta)
        //{   
           
        //    if (File.Exists(cesta))
        //    {
        //        //(App, Doc) = ExcelApp.DokumetExcel(cesta);
        //        DokumetExcel(cesta);
        //        //if (Doc == null) return (App, Doc, null);
        //        //if (Doc == null) return (App, Doc, null);
        //        if (Doc == null) return;
        //        //Nastavení listu
        //        GetSheet("Seznam Elektro");
        //        //if (Doc == null) return (App, Doc, Xls);
        //        if (Doc == null) return;
        //    }
        //    else
        //    {
        //        //VytvorNovyDokument();
        //        //(App, Doc) = VytvorNovyDokument();
        //        PridatNovyList("Seznam Elektro");
        //    }
        //    Xls.Activate();
        //    //return (App, Doc, Xls);
        //    return;
        //}
    }
}
