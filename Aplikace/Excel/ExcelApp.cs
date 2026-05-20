using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ComponentModel.DataAnnotations;
using Aplikace.Sdilene;
using Aplikace.Tridy;

namespace Aplikace.Excel
{
    /// <summary>
    /// Obalovací třída (Adapter) pro ClosedXML list (IXLWorksheet), která simuluje chování původního COM rozhraní Microsoft.Office.Interop.Excel.Worksheet.
    /// Zajišťuje zpětnou kompatibilitu a umožňuje volání vlastností jako Range, Cells a Name bez změny původního kódu.
    /// </summary>
    /// <remarks>
    /// Inicializuje novou instanci třídy ExcelWorksheetWrapper obalující zadaný ClosedXML list.
    /// </remarks>
    public class ExcelWorksheetWrapper(IXLWorksheet ws) {
        private readonly IXLWorksheet _ws = ws;

        /// <summary>
        /// Získá podkladový objekt ClosedXML IXLWorksheet.
        /// </summary>
        public IXLWorksheet Worksheet => _ws;

        /// <summary>
        /// Získá název listu.
        /// </summary>
        public string Name => _ws.Name;

        /// <summary>
        /// Zprostředkuje přístup k rozsahům buněk přes Interop-like syntaxi Range["A1"].
        /// </summary>
        public ExcelRangeWrapper Range => new(_ws);

        /// <summary>
        /// Zprostředkuje přístup k buňkám přes indexovaný přístup Cells[radek, sloupec].
        /// </summary>
        public ExcelCellsWrapper Cells => new(_ws);

        /// <summary>
        /// Aktivuje aktuální list (v ClosedXML je to prázdná operace, zachovaná pro kompatibilitu).
        /// </summary>
        public static void Activate()
        {
            // Dummy implementace pro zachování kompatibility s Interopem
        }
    }

    /// <summary>
    /// Zajišťuje indexovaný přístup k buňkám a rozsahům pomocí textové adresy (např. Range["A1"]) 
    /// nebo dvou rohových buněk (např. Range[cell1, cell2]).
    /// </summary>
    /// <remarks>
    /// Inicializuje novou instanci indexeru rozsahů.
    /// </remarks>
    public class ExcelRangeWrapper(IXLWorksheet ws) {
        private readonly IXLWorksheet _ws = ws;

        /// <summary>
        /// Získá nebo nastaví buňku na základě textové adresy (např. Range["A2"]).
        /// </summary>
        public ExcelCellWrapper this[string address]
        {
            get => new(_ws.Cell(address));
            set
            {
                var cell = _ws.Cell(address);
                if (value != null)
                {
                    cell.Value = XLCellValue.FromObject(value.Value);
                }
            }
        }

        /// <summary>
        /// Získá rozsah buněk na základě dvou rohových buněk (např. Range[Xls.Cells[1,1], Xls.Cells[2,2]]).
        /// </summary>
        public ExcelCellWrapper this[object cell1, object cell2]
        {
            get
            {
                //porovnání vzorů
                if (cell1 is ExcelCellWrapper c1 && cell2 is ExcelCellWrapper c2)
                {
                    return new ExcelCellWrapper(_ws.Range(c1.Cell, c2.Cell));
                }
                return new ExcelCellWrapper(_ws.Cell(1, 1));
            }
        }
    }

    /// <summary>
    /// Zajišťuje indexovaný přístup k jednotlivým buňkám pomocí souřadnic řádku a sloupce (např. Cells[3, 4]).
    /// </summary>
    /// <remarks>
    /// Inicializuje novou instanci indexeru buněk.
    /// </remarks>
    public class ExcelCellsWrapper(IXLWorksheet ws)
    {
        private readonly IXLWorksheet _ws = ws;

        /// <summary>
        /// Získá obalenou buňku na zadaném řádku a sloupci (indexováno od 1).
        /// </summary>
        public ExcelCellWrapper this[int row, int col]
        {
            get => new(_ws.Cell(row, col));
        }
    }

    /// <summary>
    /// Obalovací třída pro buňku (IXLCell) nebo rozsah (IXLRange) ClosedXML, která simuluje chování COM objektu Range.
    /// Poskytuje vlastnosti pro přístup k hodnotám, vzorcům a lokálním vzorcům.
    /// </summary>
    public class ExcelCellWrapper
    {
        private readonly IXLRange _range;
        private readonly IXLCell _cell;

        /// <summary>
        /// Inicializuje obal pro jedinou ClosedXML buňku.
        /// </summary>
        public ExcelCellWrapper(IXLCell cell)
        {
            _cell = cell;
        }

        /// <summary>
        /// Inicializuje obal pro ClosedXML rozsah buněk.
        /// </summary>
        public ExcelCellWrapper(IXLRange range)
        {
            _range = range;
            _cell = range.FirstCell();
        }

        /// <summary>
        /// Získá podkladovou buňku.
        /// </summary>
        public IXLCell Cell => _cell;

        /// <summary>
        /// Získá podkladový rozsah.
        /// </summary>
        public IXLRange Range => _range;

        /// <summary>
        /// Získá nebo nastaví hodnotu buňky s automatickou konverzí datových typů.
        /// </summary>
        public object Value
        {
            get => _cell?.Value ?? "";
            set
            {
                if (value == null)
                {
                    _cell?.Clear();
                    return;
                }

                if (value is string s)
                {
                    if (_cell != null) _cell.Value = s;
                    else _range?.Value = s;
                }
                else if (value is double d)
                {
                    if (_cell != null) _cell.Value = d;
                    else _range?.Value = d;
                }
                else if (value is int val)
                {
                    if (_cell != null) _cell.Value = val;
                    else _range?.Value = val;
                }
                else if (value is bool b)
                {
                    if (_cell != null) _cell.Value = b;
                    else _range?.Value = b;
                }
                else
                {
                    string str = value.ToString() ?? "";
                    if (_cell != null) _cell.Value = str;
                    else _range?.Value = str;
                }
            }
        }

        /// <summary>
        /// Získá nebo nastaví hodnotu buňky (alternativní název vlastnosti pro kompatibilitu).
        /// </summary>
        public object Value2
        {
            get => Value;
            set => Value = value;
        }

        /// <summary>
        /// Získá nebo nastaví vzorec v buňce ve formátu A1 (např. "=SUM(D3:D10)").
        /// </summary>
        public string Formula
        {
            get => _cell?.FormulaA1 ?? "";
            set
            {
                _cell?.FormulaA1 = value;
            }
        }

        /// <summary>
        /// Získá nebo nastaví lokalizovaný vzorec buňky (v ClosedXML mapováno na standardní FormulaA1).
        /// </summary>
        public string FormulaLocal
        {
            get => Formula;
            set => Formula = value;
        }
    }

    public class ExcelApp
    {
        public XLWorkbook Doc { get; set; }
        public ExcelWorksheetWrapper Xls { get; set; }
        public int Process { get; set; }
        public object App { get; set; } // Dummy property to prevent compilation errors in other files
        private int record = 1;

        public ExcelApp() : this("")
        { }

        public ExcelApp(string Cesta)
        {
            if (!string.IsNullOrEmpty(Cesta) && File.Exists(Cesta))
            {
                Console.WriteLine("Otevřít dokument Excel.");
                Doc = new XLWorkbook(Cesta);
                Xls = new ExcelWorksheetWrapper(Doc.Worksheet(1));
                return;
            }

            Console.WriteLine("Vytvořen prázdný dokument Excel.");
            Doc = new XLWorkbook();
            var ws = Doc.Worksheets.Add("List1");
            Xls = new ExcelWorksheetWrapper(ws);
            Console.WriteLine("\nVytvořený dokument nastaven Aktivní.");
        }

        public void GetSheet(string Nazev)
        {
            if (Doc == null) return;
            if (Doc.TryGetWorksheet(Nazev, out var ws))
            {
                Xls = new ExcelWorksheetWrapper(ws);
                Console.WriteLine($"List {ws.Name} - Nastaven");
                return;
            }
            var newWs = Doc.Worksheets.Add(Nazev);
            Xls = new ExcelWorksheetWrapper(newWs);
            Console.WriteLine($"List {Nazev} - Přidán");
        }

        public void DokumetExcel(string Cesta)
        {
            Console.Write("\nKontrolaOtevenehoNeboOtevreniSobroruExel - OK");
            KontrolaOtevenehoNeboOtevreniSobroruExel(Cesta);
        }

        public void KontrolaOtevenehoNeboOtevreniSobroruExel(string Cesta)
        {
            Console.Write("\nMetoda Kontrola Oteveneho Nebo Otevreni Sobroru Exel");
            Console.Write("\nCesta" + Cesta.ToLowerInvariant());
            if (File.Exists(Cesta))
            {
                Console.Write("\nSoubor není otevřen kontrola ");
                Doc = new XLWorkbook(Cesta);
                Xls = new ExcelWorksheetWrapper(Doc.Worksheet(1));
            }
            else
            {
                Doc = new XLWorkbook();
                var ws = Doc.Worksheets.Add("List1");
                Xls = new ExcelWorksheetWrapper(ws);
            }
        }

        public List<List<string>> ExelLoadTable(string cesta, string zalozka, int Radek, int[] CteniSloupcu)
        {
            if (!File.Exists(cesta)) return [];

            DokumetExcel(cesta);
            if (Xls == null) return [];
            Console.Write("\nDokument excel - Otevřen");

            GetSheet(zalozka);
            if (Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.Write("\nSheet=" + Xls.Name);

            var Pole = new List<List<string>>();
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;
            Console.Write("\nZal.Rows.Count=" + rowCount);

            for (int i = Radek; i <= rowCount; i++)
            {
                var cteniPole = new List<string>();
                foreach (var item in CteniSloupcu)
                {
                    var cell = ws.Cell(i, item);
                    string xxx = cell.GetString();
                    cteniPole.Add(xxx);
                }
                if (cteniPole.Count > 1 && !string.IsNullOrEmpty(cteniPole[1]) && cteniPole[1] != "0")
                {
                    Pole.Add(cteniPole);
                    Console.Write("\nRadek=" + i.ToString() + "\t" + cteniPole[0]);
                }
                if (i > 100 && Pole.Count > 0 && Pole.Last().First().Length < 2) break;
            }
            Console.Write("\nUkončení Excel");
            ExcelQuit(cesta);
            Console.Write("\nUkončení Excel");
            return Pole;
        }

        public List<Zarizeni> ExelTable(int Radek, string Tabulka, IDictionary<int, string> dir)
        {
            GetSheet(Tabulka);
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            int pocet = 1;
            var Pole = new List<Zarizeni>();
            int colCount = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            Console.WriteLine($"[Rows.Col]=[{rowCount},{colCount}]");

            for (int i = Radek; i <= rowCount; i++)
            {
                var jeden = new Zarizeni();
                bool Prerusit = true;

                //OMEZENÍ NAČÁTÁNÍ RADKU
                //string tagstr = ws.Cell(i, 5).GetString();
                //if (!new[] { "M", "MOB", "MOP" }.Contains(tagstr))
                //{
                //    Prerusit = false;
                //    continue;
                //}

                foreach (var j in dir.Keys.ToArray())
                {
                    var cell = ws.Cell(i, j);
                    if (cell.IsMerged())
                    {
                        Console.WriteLine("Buňka je součástí sloučených buněk.");
                        break;
                    }
                    string xxx = cell.GetString().Replace('\n', ' ');

                    if (string.IsNullOrEmpty(xxx) || xxx == "0")
                        continue;

                    //zdroj je int
                    if (int.TryParse(xxx, out int intVal) )
                        //proměná pro vložení je int
                        if (jeden[dir[j]].GetType() == typeof(int))
                            jeden[dir[j]] = intVal;
                        else
                            //proměná není int ale zdroj je int, vloží se jako string
                            jeden[dir[j]] = xxx;
                    else
                        //zdroj není int, vloží se jako string
                        jeden[dir[j]] = xxx;
                }
                

                jeden.Apid = ExcelLoad.Apid();
                jeden.Id = pocet;

                if (jeden.Pocet > 1)
                {
                    var deleni = jeden.Tag?.Split('\n').ToList() ?? [];
                    foreach (var item in deleni)
                    {
                        var json = System.Text.Json.JsonSerializer.Serialize(jeden);
                        var kopie = System.Text.Json.JsonSerializer.Deserialize<Zarizeni>(json)!;
                        kopie.Apid = ExcelLoad.Apid();
                        kopie.Pocet = 1;
                        kopie.Tag = item.Trim();
                        Pole.Add(kopie);
                        Console.WriteLine($"Tag {kopie.Tag}");
                    }
                }
                else if (Prerusit)
                {
                    Pole.Add(jeden);
                }
                Console.WriteLine($"Radek {pocet++} - přídán");
            }
            Console.WriteLine("Zavřít sešit Excel");
            return Pole;
        }

        public List<Vykres> ExelTableVykresy(int Radek, string Tabulka, IDictionary<int, string> dir)
        {
            GetSheet(Tabulka);
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            int pocet = 1;
            var Pole = new List<Vykres>();
            int colCount = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            Console.WriteLine($"[Rows.Col]=[{rowCount},{colCount}]");

            for (int i = Radek; i <= rowCount; i++)
            {
                var jeden = new Vykres();
                foreach (var j in dir.Keys.ToArray())
                {
                    var cell = ws.Cell(i, j);
                    if (cell.IsMerged())
                    {
                        Console.WriteLine("Buňka je součástí sloučených buněk.");
                        break;
                    }
                    string xxx = cell.GetString();
                    jeden[dir[j]] = xxx.Trim();
                }
                if (!string.IsNullOrEmpty(jeden.Nazev))
                    Pole.Add(jeden);
                Console.WriteLine($"Radek {pocet++} - přídán");
            }
            return Pole;
        }

        public void ClassToExcel<T>(int Row, List<T> Pole, IDictionary<int, string> Sloupce)
        {
            var properties = typeof(T).GetProperties().ToDictionary(p => p.Name);

            var dirFiltered = Sloupce
                .Where(kvp => properties.ContainsKey(kvp.Value))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            var ws = Xls.Worksheet;

            foreach (var item in Pole)
            {
                foreach (var kvp in dirFiltered)
                {
                    var cell = ws.Cell(Row, kvp.Key);
                    var prop = properties[kvp.Value];
                    var rawValue = prop.GetValue(item);
                    string value = rawValue?.ToString() ?? "";

                    if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double cislo))
                    {
                        if (prop.Name == "Delka") cislo /= 1000;
                        cell.Value = cislo;
                        cell.Style.NumberFormat.Format = "#,##0.00";
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else
                    {
                        cell.Value = value;
                    }
                }
                Row++;
            }

            foreach (var colKey in dirFiltered.Keys)
            {
                ws.Column(colKey).AdjustToContents();
            }
        }

        public List<Zarizeni> ExelLoadTableTrida(string cesta, string zalozka, int Radek, int[] CteniSloupcu, string[] TextPole)
        {
            if (!File.Exists(cesta)) return [];

            DokumetExcel(cesta);
            if (Xls == null) return [];
            Console.Write("\nDokument excel - Otevřen");

            GetSheet(zalozka);
            if (Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.Write("\nSheet=" + Xls.Name);

            var Pole = new List<Zarizeni>();
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;
            Console.Write("\nZal.Rows.Count=" + rowCount);

            for (int i = Radek; i <= rowCount; i++)
            {
                var obj = new Zarizeni();
                int x = 0;
                foreach (var item in CteniSloupcu)
                {
                    var cell = ws.Cell(i, item);
                    string xxx = cell.GetString();
                    if (!string.IsNullOrEmpty(xxx))
                    {
                        obj[TextPole[x++]] = xxx;
                    }
                }
                Pole.Add(obj);
            }
            ExcelQuit(cesta);
            return Pole;
        }

        public void ExcelSaveJeden(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
        {
            if (!File.Exists(cesta)) return;

            DokumetExcel(cesta);
            if (Xls == null) return;
            Console.Write("\nDokument excel - Otevřen");

            GetSheet(zalozka);
            if (Xls == null) { Console.Write("\nChyba KONEC"); return; }
            Console.Write("\nSheet=" + Xls.Name);

            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 7; i <= rowCount; i++)
            {
                var cteniPole = new List<string>();
                foreach (var item in SloupceCteni)
                {
                    var cell = ws.Cell(i, item);
                    string xxx = cell.GetString();
                    if (!string.IsNullOrEmpty(xxx))
                        cteniPole.Add(xxx);
                }

                var Shoda = Vstup.FirstOrDefault(x => x.FirstOrDefault() == cteniPole.FirstOrDefault());

                if (Shoda != null)
                {
                    Console.Write("\nShoda buňky " + i + " = " + Shoda.First());
                    ws.Cell(i, SloupceZapisu.First()).Value = Shoda.First();
                    ws.Cell(i, SloupceZapisu.Last()).Value = Shoda[8] + " " + Shoda.Last();
                }
                else
                {
                    ws.Cell(i, SloupceZapisu.First()).Value = "Nenalezeno";
                }

                Console.Write("\nShoda buňky " + i);
                if (i > 500) break;
            }
            Console.Write("\nUkončení Excel");
            Doc.Save();
            Console.Write("\nSave OK");
            ExcelQuit(cesta);
            Console.Write("\nUkončení Excel");
        }

        public void ExcelSaveSloupec(string cesta, int[] SloupceZapisu, string zalozka, int[] SloupceCteni, List<List<string>> Vstup)
        {
            string cesta1 = @"C:\VisualStudio\Parametr\AplikacePomoc\Motory\Motory500V.xlsx";
            var PouzitProTabulku = new int[] { 1, 2, 3 };
            var Motory500 = ExcelLoad.LoadDataExcel(cesta1, PouzitProTabulku, "Motory500V", 2);
            Motory500.Vypis();

            if (!File.Exists(cesta)) return;

            DokumetExcel(cesta);
            if (Xls == null) return;
            Console.Write("\nDokument excel - Otevřen");

            GetSheet(zalozka);
            if (Xls == null) return;
            Console.Write("\nSheet=" + Xls.Name);

            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 7; i <= rowCount; i++)
            {
                var cteniPole = new List<string>();
                foreach (var item in SloupceCteni)
                {
                    var cell = ws.Cell(i, item);
                    string xxx = cell.GetString();
                    if (!string.IsNullOrEmpty(xxx))
                        cteniPole.Add(xxx);
                }

                var Shoda = Vstup.FirstOrDefault(x => x.FirstOrDefault() == cteniPole.FirstOrDefault());

                if (Shoda != null)
                {
                    Console.Write("\nShoda buňky " + i + " = " + Shoda.First());

                    if (cteniPole.Count > 1 && double.TryParse(cteniPole[1], out double Prikon))
                    {
                        var Informace = Motory500.FirstOrDefault(x => Convert.ToDouble(x[0]) == Prikon)?[1];
                        if (double.TryParse(Informace, out double Proud))
                        {
                            ws.Cell(i, SloupceZapisu.First()).Value = Proud;
                        }
                    }

                    ws.Cell(i, SloupceZapisu[1]).Value = Shoda[8];
                    ws.Cell(i, SloupceZapisu[2]).Value = Shoda[9];

                    if (double.TryParse(Shoda[4], out double delka))
                        ws.Cell(i, SloupceZapisu[3]).Value = delka;
                    else
                        ws.Cell(i, SloupceZapisu[3]).Value = Shoda[4];

                    if (double.TryParse(Shoda[5], out double AWG))
                        ws.Cell(i, SloupceZapisu[5]).Value = AWG;
                    else
                        ws.Cell(i, SloupceZapisu[5]).Value = Shoda[5];

                    if (double.TryParse(Shoda[10], out double mm2))
                        ws.Cell(i, SloupceZapisu[4]).Value = mm2;
                    else
                        ws.Cell(i, SloupceZapisu[4]).Value = "";
                }
                else
                {
                    ws.Cell(i, SloupceZapisu.First()).Value = "Nenalezeno";
                }

                Console.Write("\nShoda buňky " + i);
                if (i > 500) break;
            }
            ExcelQuit(cesta);
        }

        public void ExcelSaveT<T>(T[] pole, string Nazev)
        {
            string ClassName = typeof(T).Name;
            Console.WriteLine(ClassName);

            var TridaPole = pole.GetType();
            Console.WriteLine(TridaPole.Name);

            var Sloupce = typeof(T).GetProperties();
            foreach (var item in Sloupce)
                Console.WriteLine(item.Name);

            int row = 1; int col = 1;
            var ws = Xls.Worksheet;
            ws.Cell(row, col).Value = Nazev;
            row++;
            foreach (var item in Sloupce)
            {
                DisplayAttribute displayAttribute = item.GetCustomAttributes(typeof(DisplayAttribute), false).Cast<DisplayAttribute>().FirstOrDefault();
                ws.Cell(row, col).Value = displayAttribute != null ? displayAttribute.Name : item.Name.ToUpper();
                col++;
            }

            col = 1;
            row++;
            foreach (var item in pole)
            {
                foreach (var Property in Sloupce)
                {
                    Console.WriteLine(Property.PropertyType.ToString());

                    if (Property.PropertyType == typeof(string))
                    {
                        var value = item?.GetType().GetProperty(Property.Name)?.GetValue(item)?.ToString();
                        ws.Cell(row, col).Value = value;
                        col++;
                    }
                }
                col = 1; row++;
            }
        }

        public void NadpisMIlan()
        {
            string Nad = @"    |     |   |     |     |                                        |  |KAPACITA        |                        |        |        |      |EL.  |        ";
            int col = 1;
            int row = 1;
            var ws = Xls.Worksheet;
            foreach (var item in Nad.Split('|'))
            {
                ws.Cell(row, col++).Value = item;
            }
            row++; col = 1;
            Nad = "GUID|IO/SO|NO |PS   |TAG  |NÁZEV                                   |KS|NOSTNOST        |MEDIUM                  |OBJEM   |PRŮTOK  |HMOTN.|PŘÍK.|POZNÁMKA";
            foreach (var item in Nad.Split('|'))
            {
                ws.Cell(row, col++).Value = item;
            }

            var range = ws.Range(1, 1, 2, col - 1);
            range.Style.Alignment.WrapText = false;
            NadpisSet(range);
        }

        public void ExcelSave(Item[] pole)
        {
            NadpisMIlan();
            int col = 1;
            int row = 3;
            Tisk(pole, ref row, col);

            for (int i = 1; i <= 20; i++)
                Xls.Worksheet.Column(i).AdjustToContents();
        }

        public int Tisk(Item[] pole, ref int row, int col)
        {
            var ws = Xls.Worksheet;
            foreach (var item in pole)
            {
                ws.Cell(row, col++).Value = item.Id.ToString();
                ws.Cell(row, col++).Value = item.Cunit.Pfx + " " + item.Cunit.Num;
                ws.Cell(row, col++).Value = (record++).ToString();
                ws.Cell(row, col++).Value = item.Munit.Pfx + " " + item.Munit.Num;
                ws.Cell(row, col++).Value = item.Tag;
                ws.Cell(row, col++).Value = item.Name;
                ws.Cell(row, col++).Value = item.Pcs;

                ws.Cell(row, col + 4).Value = item.Mass;
                ws.Cell(row, col + 5).Value = item.Power;
                ws.Cell(row, col + 6).Value = item.Note;

                if (item.Fluid.Count > 0)
                {
                    if (item.Fluid.Count > 1) row++;
                    foreach (var item2 in item.Fluid)
                    {
                        ws.Cell(row, col).Value = item2.Parameter.Value.ToString() + " " + item2.Parameter.Unit;
                        ws.Cell(row, col + 1).Value = item2.Fluid;
                        ws.Cell(row, col + 2).Value = item2.Volume;
                        ws.Cell(row, col + 3).Value = item2.Flowrate;
                        row++;
                    }
                    col += 4; row--;
                }
                else
                    col += 4;

                var range = ws.Range(row, 1, row, col);
                range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                if (record % 2 == 1)
                    range.Style.Fill.BackgroundColor = XLColor.LightGray;

                if (item.Subitem.Count > 0)
                {
                    row++; col = 1;
                    Tisk([.. item.Subitem], ref row, col);
                }
                else
                {
                    row++; col = 1;
                }
            }
            return row;
        }

        public static void NadpisSet(IXLRange range)
        {
            range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            range.Style.Border.LeftBorderColor = XLColor.Black;
            range.Style.Border.RightBorderColor = XLColor.Black;
            range.Style.Border.TopBorderColor = XLColor.Black;
            range.Style.Border.BottomBorderColor = XLColor.Black;

            SetFontRed(range.Style.Font);

            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            range.Style.Fill.BackgroundColor = XLColor.LightBlue;

            foreach (var col in range.Worksheet.Columns(range.RangeAddress.FirstAddress.ColumnNumber, range.RangeAddress.LastAddress.ColumnNumber))
            {
                col.AdjustToContents();
            }

            foreach (var r in range.Worksheet.Rows(range.RangeAddress.FirstAddress.RowNumber, range.RangeAddress.LastAddress.RowNumber))
            {
                r.AdjustToContents();
            }
        }

        public IXLRange Nadpisy(Nadpis[] data)
        {
            int col = 1;
            var ws = Xls.Worksheet;
            foreach (var item in data)
            {
                ws.Cell(1, col).Value = item.Name;
                ws.Cell(2, col++).Value = item.Jednotky;
            }

            var range = ws.Range(1, 1, 2, col - 1);
            range.Style.Alignment.WrapText = true;
            return range;
        }

        public IXLRange Nadpisy(IDictionary<int, string> Dir)
        {
            int col = 1;
            var ws = Xls.Worksheet;
            var properties = new List<PropertyInfo>();

            Type currentType = typeof(Slaboproudy);
            properties.AddRange(currentType.GetProperties(BindingFlags.Public | BindingFlags.Instance));

            var ppp = properties.ToDictionary(p => p.Name);

            foreach (var kvp in Dir)
            {
                if (!ppp.ContainsKey(kvp.Value)) continue;
                var prop = ppp[kvp.Value];

                if (prop == null) continue;

                var displayAttr = prop.GetCustomAttribute<DisplayAttribute>();
                string displayName = displayAttr?.Name ?? prop.Name;

                var jednotkyAttr = prop.GetCustomAttribute<JednotkyAttribute>();
                string jednotky = jednotkyAttr?.Text ?? "";

                ws.Cell(1, kvp.Key).Value = displayName;
                ws.Cell(2, kvp.Key).Value = jednotky;
                col++;
            }

            var range = ws.Range(1, 1, 2, col - 1);
            range.Style.Alignment.WrapText = false;
            return range;
        }

        public void ExcelSaveList(List<List<string>> Vstup)
        {
            int row = 2; int col = 1;
            var ws = Xls.Worksheet;

            var Kontrola = ws.Cell(row + 1, col);
            if (!Kontrola.IsEmpty())
            {
                Console.WriteLine("Přepsat");
                if (Console.ReadKey().Key != ConsoleKey.A) return;
            }

            foreach (var radek in Vstup)
            {
                row++; col = 1;
                foreach (var item in radek)
                {
                    var cell = ws.Cell(row, col++);
                    if (double.TryParse(item, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double cislo))
                        cell.Value = cislo;
                    else
                        cell.Value = item;
                }
                ws.Row(row).AdjustToContents();
            }
            ws.Columns().AdjustToContents();
        }

        public void ExcelSaveClass(List<Zarizeni> Vstup)
        {
            int row = 2;
            var ws = Xls.Worksheet;

            foreach (var radek in Vstup)
            {
                row++;
                for (int col = 1; col <= 15; col++)
                {
                    var cell = ws.Cell(row, col);
                    switch (col)
                    {
                        case 1:
                            cell.Value = radek.Tag;
                            break;
                        case 2:
                            cell.Value = radek.Pid;
                            break;
                        case 3:
                            cell.Value = radek.Popis;
                            break;
                        case 4:
                            cell.Value = radek.Prikon;
                            break;
                        case 5:
                            cell.Value = radek.BalenaJednotka;
                            break;
                        case 6:
                            cell.Value = radek.Menic;
                            break;
                        case 7:
                            cell.Value = radek.Proud;
                            break;
                        case 8:
                            cell.Value = radek.HP;
                            break;
                        case 9:
                            if (double.TryParse(radek.Proud, out double proud1))
                                cell.Value = (proud1 * 500 / 480).ToString();
                            break;
                        case 10:
                            cell.Value = radek.PrurezMM2;
                            break;
                        case 12:
                            cell.Value = radek.Delka;
                            break;
                        case 14:
                            cell.Value = radek.Rozvadec;
                            break;
                        case 15:
                            cell.Value = radek.RozvadecCislo;
                            break;
                        default:
                            break;
                    }
                }
                ws.Row(row).AdjustToContents();
            }
            ws.Columns().AdjustToContents();
        }

        public void ExcelSaveProud(List<List<string>> Vstup)
        {
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 3; i <= rowCount; i++)
            {
                var cell = ws.Cell(i, 4);
                string xxx = cell.GetString();
                if (double.TryParse(xxx, out double cislo))
                {
                    var Informace = Vstup.FirstOrDefault(x => Convert.ToDouble(x[0]) == cislo)?[1];
                    if (double.TryParse(Informace, out double Proud))
                    {
                        ws.Cell(i, 7).Value = Proud;
                    }
                }

                if (cell.IsEmpty() && i > 100)
                    break;
            }
        }

        public void ExcelSaveVzorce(int Pocet)
        {
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 3; i <= rowCount; i++)
            {
                ws.Cell(i, 8).FormulaA1 = $"=D{i}*1.341022";
                ws.Cell(i, 9).FormulaA1 = $"=G{i}*500/480";
                ws.Cell(i, 13).FormulaA1 = $"=L{i}*3.280839895";

                if (i > Pocet)
                    break;
            }
        }

        public System.Data.DataTable GetTable(int rowNadpis, int[] sloupec)
        {
            var Table = new System.Data.DataTable("Tabulka");
            var ws = Xls.Worksheet;
            int usedRows = ws.LastRowUsed()?.RowNumber() ?? 0;

            foreach (var i in sloupec)
            {
                string cellValue = ws.Cell(rowNadpis, i).GetString().Trim();
                Console.Write("\nRadek=" + rowNadpis + ", Sloupec=" + i + ", nadpis=" + cellValue);
                Table.Columns.Add(string.IsNullOrEmpty(cellValue) ? i.ToString() : cellValue, typeof(string));
            }

            int t = 0;
            for (int row = rowNadpis + 1; row <= usedRows; row++)
            {
                var Pole = new List<string>();
                var rada = Table.NewRow();
                int colpomoc = 0;
                string text = string.Empty;

                foreach (var col in sloupec)
                {
                    var cell = ws.Cell(row, col);
                    string cteni = cell.GetString();
                    Pole.Add(cteni);
                    text += cteni;
                    rada[colpomoc++] = cteni;
                    Console.Write("\ncteni " + cteni);
                }

                Table.Rows.Add(rada);

                Console.Write("\nDelka" + text.Length);
                if (text.Length < 4) return Table;
                if (t++ > 1000) return Table;
            }
            return Table;
        }

        public static bool ExcelKontrolaInstalace()
        {
            return true;
        }

        public bool ExcelQuit(string cesta)
        {
            Console.Write("\nUkončení Excel, ");
            if (Doc != null)
            {
                Doc.SaveAs(cesta);
                Console.Write("\nSave OK");
            }
            return true;
        }

        public void ExcelSaveKabel(List<List<string>> Vstup)
        {
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 2; i <= rowCount; i++)
            {
                string xxx = ws.Cell(i, 1).GetString();

                var Informace = Vstup.FirstOrDefault(x => x[0] == xxx);

                if (Informace != null)
                {
                    if (double.TryParse(Informace[4], out double delka))
                    {
                        ws.Cell(i, 12).Value = delka;
                    }

                    ws.Cell(i, 11).Value = Informace[5];

                    if (double.TryParse(Informace[10], out double mm2))
                    {
                        ws.Cell(i, 10).Value = mm2;
                    }
                }

                if (ws.Cell(i, 1).IsEmpty() && i > 100)
                    break;
            }
        }

        public static void ExcelSaveRozvadec(ExcelWorksheetWrapper ListExcel, List<List<string>> Vstup)
        {
            var ws = ListExcel.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 2; i <= rowCount; i++)
            {
                string xxx = ws.Cell(i, 1).GetString();

                var Informace = Vstup.FirstOrDefault(x => x[0] == xxx);

                if (Informace != null)
                {
                    ws.Cell(i, 14).Value = Informace[8];

                    if (double.TryParse(Informace[9], out double cislo))
                    {
                        ws.Cell(i, 15).Value = cislo;
                    }
                }
            }
        }

        public List<List<string>> ExcelLoadWorksheet(int[] pouzitProTabulku)
        {
            var Data = new List<List<string>>();
            var ws = Xls.Worksheet;
            int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 3; i <= rowCount; i++)
            {
                var Radek = new List<string>();
                string Cteni = "";
                foreach (var item in pouzitProTabulku)
                {
                    var cell = ws.Cell(i, item);
                    Cteni = cell.GetString();
                    Radek.Add(Cteni);
                }
                Data.Add(Radek);

                if (string.IsNullOrEmpty(Cteni) && i > 100)
                    break;
            }
            return Data;
        }

        public void KabelyToExcel(List<List<string>> data, int Row)
        {
            Row--;
            int j = 1;
            var ws = Xls.Worksheet;
            foreach (var radek in data)
            {
                Console.WriteLine("Radek " + Row);
                Row++; j = 1;
                foreach (var item in radek)
                {
                    var cell = ws.Cell(Row, j++);
                    if (double.TryParse(item, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double cislo))
                    {
                        cell.Value = cislo;
                    }
                    else
                    {
                        cell.Value = item;
                    }
                }
            }
            for (int i = 1; i < j; i++)
                ws.Column(i).AdjustToContents();
        }

        public void KabelyToExcel(List<string> data, int Row)
        {
            Row--;
            int j = 1;
            var ws = Xls.Worksheet;
            foreach (var item in data)
            {
                Console.WriteLine("Radek " + Row);
                Row++; j = 1;
                var cell = ws.Cell(Row, j++);
                cell.Value = item;
            }
            for (int i = 1; i < j; i++)
                ws.Column(i).AdjustToContents();
        }

        public void ExcelSaveNadpis<T>(List<T> Ramecek)
        {
            var VelikostTabulky = Ramecek.Count;
            Nadpis("A1:D1", "Označeni", VelikostTabulky);
            Nadpis("E1:H1", "Kabel", VelikostTabulky);
            Nadpis("I1:I1", "Zařízení", VelikostTabulky);
            Nadpis("J1:M1", "Odkud", VelikostTabulky);
            Nadpis("N1:N1", "", VelikostTabulky);
            Nadpis("O1:R1", "Kam", VelikostTabulky);
            Nadpis("S1:S1", "Delka", VelikostTabulky);

            Xls.Range["A2"].Value = "Tag";
            Xls.Range["B2"].Value = "MCC";
            Xls.Range["D2"].Value = "Číslo";

            Xls.Range["E2"].Value = "Kabel";
            Xls.Range["F2"].Value = "Vodic";
            Xls.Range["G2"].Value = "[mm2]";

            Xls.Range["J2"].Value = "Tag";
            Xls.Range["K2"].Value = "MCC";
            Xls.Range["M2"].Value = "Svorka";

            Xls.Range["O2"].Value = "Tag";
            Xls.Range["P2"].Value = "Patro";
            Xls.Range["R2"].Value = "Svorka";

            Xls.Range["S2"].Value = "[m]";
        }

        public void ExcelSaveNadpisEn<T>(List<T> Ramecek)
        {
            var VelikostTabulky = Ramecek.Count;
            Nadpis("A1:D1", "TAG", VelikostTabulky);
            Nadpis("E1:H1", "CABLE", VelikostTabulky);
            Nadpis("I1:I1", "TYPE", VelikostTabulky);
            Nadpis("J1:M1", "FROM", VelikostTabulky);
            Nadpis("N1:N1", "", VelikostTabulky);
            Nadpis("O1:R1", "TO", VelikostTabulky);
            Nadpis("S1:S1", "LENGHT", VelikostTabulky);

            Xls.Range["A2"].Value = "TAG";
            Xls.Range["B2"].Value = "MCC";
            Xls.Range["D2"].Value = "NUMBER";

            Xls.Range["E2"].Value = "CABLE";
            Xls.Range["F2"].Value = "CONDUCTOR";
            Xls.Range["G2"].Value = "[mm2]";

            Xls.Range["J2"].Value = "TAG";
            Xls.Range["K2"].Value = "MCC";
            Xls.Range["M2"].Value = "CLAMP";

            Xls.Range["O2"].Value = "TAG";
            Xls.Range["P2"].Value = "FLOOR";
            Xls.Range["R2"].Value = "CLAMP";

            Xls.Range["S2"].Value = "[m]";
        }

        public void Nadpis(string pole, string Text)
        {
            Nadpis(pole, Text, 1);
        }

        public void Nadpis(string pole, string Text, int VelikostTabulky)
        {
            var ws = Xls.Worksheet;
            var range = ws.Range(pole);
            if (range.Cells().Count() > 1)
                range.Merge();

            range.Style.Alignment.WrapText = false;
            range.Value = Text;

            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            range.Style.Border.OutsideBorderColor = XLColor.Black;

            SetFont(range.Style.Font);

            string v = string.Concat(pole[..^1], (VelikostTabulky + 2).ToString());
            var borderRange = ws.Range(v);
            Ramecek(borderRange.Style.Border);
        }

        public static void SetFont(IXLFont Fonty)
        {
            Fonty.Bold = true;
            Fonty.FontSize = 14;
            Fonty.FontName = "Arial";
        }

        public static void SetFontRed(IXLFont Fonty)
        {
            Fonty.FontColor = XLColor.Red;
            SetFont(Fonty);
        }

        public static void Ramecek(IXLBorder borders)
        {
            borders.TopBorder = XLBorderStyleValues.Thin;
            borders.BottomBorder = XLBorderStyleValues.Thin;
            borders.LeftBorder = XLBorderStyleValues.Thin;
            borders.RightBorder = XLBorderStyleValues.Thin;
        }
    }
}
