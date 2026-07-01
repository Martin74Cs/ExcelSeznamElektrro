#nullable disable

using ExcelGenerateSeznam;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Knihovna.KabelyXls {
    /// <summary>
    /// Generátor Excelu, který provádí úpravy v šabloně pomocí přímé manipulace s OpenXML balíčkem.
    /// Nepoužívá žádné externí knihovny pro práci s Excelem.
    /// </summary>
    public class ExcelGenerator
    {
        private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        /// <summary>
        /// Vygeneruje nový Excel soubor ze šablony na základě předaných dat.
        /// </summary>
        /// <param name="sablonaCesta">Cesta k existující vzorové šabloně XLSX.</param>
        /// <param name="vystupCesta">Cesta pro uložení nově generovaného souboru.</param>
        /// <param name="coverData">Data pro vyplnění úvodního listu Cover.</param>
        /// <param name="spotrebice">Seznam spotřebičů pro uložení do listu SEMIPROVOZ IVCHS.</param>
        public void Generuj(string sablonaCesta, string vystupCesta, CoverData coverData, List<Spotrebic> spotrebice)
        {
            if (!File.Exists(sablonaCesta))
            {
                throw new FileNotFoundException("Vzorová šablona nebyla nalezena.", sablonaCesta);
            }

            // Vytvoříme kopii šablony do výstupního umístění
            string adresar = Path.GetDirectoryName(vystupCesta);
            if (!string.IsNullOrEmpty(adresar) && !Directory.Exists(adresar))
            {
                Directory.CreateDirectory(adresar);
            }
            File.Copy(sablonaCesta, vystupCesta, true);

            // Otevřeme kopii pro aktualizaci XML souborů uvnitř ZIP archivu
            using ZipArchive archiv = ZipFile.Open(vystupCesta, ZipArchiveMode.Update);
            // 1. Správce sdílených řetězců (Shared Strings)
            SharedStringsManager sharedStrings = NacistSharedStrings(archiv);

            // 2. Správce sešitu (Workbook) pro vyhledání pojmenovaných oblastí
            WorkbookManager workbook = NacistWorkbook(archiv);

            // 3. Úprava listu Cover
            UpravCover(archiv, workbook, sharedStrings, coverData);

            // 4. Úprava listu seznam spotřebiči
            UpravSpotrebice(archiv, sharedStrings, spotrebice);

            // 5. Uložení sdílených řetězců zpět
            UlozitSharedStrings(archiv, sharedStrings);
        }

        private SharedStringsManager NacistSharedStrings(ZipArchive archiv)
        {
            ZipArchiveEntry entry = archiv.GetEntry("xl/sharedStrings.xml") ?? throw new InvalidOperationException("V šabloně chybí soubor sharedStrings.xml.");
            using Stream stream = entry.Open();
            return new SharedStringsManager(stream);
        }

        private void UlozitSharedStrings(ZipArchive archiv, SharedStringsManager manager)
        {
            ZipArchiveEntry entry = archiv.GetEntry("xl/sharedStrings.xml");
            using Stream stream = entry.Open();
            manager.Save(stream);
        }

        private WorkbookManager NacistWorkbook(ZipArchive archiv)
        {
            ZipArchiveEntry entry = archiv.GetEntry("xl/workbook.xml") ?? throw new InvalidOperationException("V šabloně chybí soubor workbook.xml.");
            using Stream stream = entry.Open();
            return new WorkbookManager(stream);
        }

        private void UpravCover(ZipArchive archiv, WorkbookManager workbook, SharedStringsManager sharedStrings, CoverData coverData)
        {
            ZipArchiveEntry entry = archiv.GetEntry("xl/worksheets/sheet1.xml") ?? throw new InvalidOperationException("V šabloně chybí soubor sheet1.xml (Cover)."); // Cover list
            XDocument doc;
            using (Stream stream = entry.Open())
            {
                doc = XDocument.Load(stream);
            }

            XElement sheetData = doc.Root.Element(Ns + "sheetData") ?? throw new InvalidOperationException("V sheet1.xml chybí element sheetData.");

            // Vyplnění standardních pojmenovaných oblastí
            NastavHodnotu(sheetData, workbook, sharedStrings, "ZAKAZNIK", coverData.Zakaznik);
            NastavHodnotu(sheetData, workbook, sharedStrings, "PROJEKT", coverData.Projekt);
            NastavHodnotu(sheetData, workbook, sharedStrings, "NAZEV", coverData.Nazev);
            NastavHodnotu(sheetData, workbook, sharedStrings, "_NA4", coverData.DokumentNazev);
            NastavHodnotu(sheetData, workbook, sharedStrings, "_CAS4", coverData.Technologie);
            NastavHodnotu(sheetData, workbook, sharedStrings, "CDOK1", coverData.CistyDokumentTyp);
            NastavHodnotu(sheetData, workbook, sharedStrings, "CDOK2", coverData.CistyDokumentCislo);
            NastavHodnotu(sheetData, workbook, sharedStrings, "REV", coverData.Revize);

            // Vyplnění revizí (max 6 řádků, indexováno od 1 do 6, řádky 52 až 47)
            // Zápis provádíme odspodu (první revize v seznamu půjde do _REV1, druhá do _REV2 atd.)
            for (int i = 0; i < 6; i++)
            {
                string revIdx = (i + 1).ToString();
                string revName = $"_REV{revIdx}";
                string datName = $"_DAT{revIdx}";
                string popName = $"_POP{revIdx}";
                string zpracName = $"ZPRAC{revIdx}"; // v definovaných názvech je ZPRAC1, ZPRAC2 atd. (bez podtržítka)
                string kontrolName = $"KONTROL{revIdx}";
                string schvalilName = $"SCHVALIL{revIdx}";

                // Najdeme řádek pro status revize. Status je ve sloupci D.
                // Zjistíme referenci buňky pro _REV1 a z ní odvodíme řádek.
                string revCellRef = workbook.GetCellRef(revName);
                string statusCellRef = null;
                if (!string.IsNullOrEmpty(revCellRef))
                {
                    string rowStr = new(revCellRef.Where(char.IsDigit).ToArray());
                    statusCellRef = $"D{rowStr}";
                }

                if (i < coverData.RevizeSeznam.Count)
                {
                    RevizeInfo rev = coverData.RevizeSeznam[i];
                    NastavHodnotu(sheetData, workbook, sharedStrings, revName, rev.Revize);
                    NastavHodnotu(sheetData, workbook, sharedStrings, datName, rev.Datum);
                    NastavHodnotu(sheetData, workbook, sharedStrings, popName, rev.PopisRevize);
                    NastavHodnotu(sheetData, workbook, sharedStrings, $"ZPRAC{revIdx}", rev.Zpacoval);
                    NastavHodnotu(sheetData, workbook, sharedStrings, kontrolName, rev.Kontroloval);
                    NastavHodnotu(sheetData, workbook, sharedStrings, schvalilName, rev.Schvalil);

                    // Zapíšeme status revize do sloupce D
                    if (!string.IsNullOrEmpty(statusCellRef))
                    {
                        int sIdx = sharedStrings.GetOrAdd(rev.Stat);
                        SetCellValue(sheetData, Ns, statusCellRef, rev.Stat, sIdx, null);
                    }
                }
                else
                {
                    // Vyčistíme nepoužité revizní řádky
                    NastavHodnotu(sheetData, workbook, sharedStrings, revName, "");
                    NastavHodnotu(sheetData, workbook, sharedStrings, datName, "");
                    NastavHodnotu(sheetData, workbook, sharedStrings, popName, "");
                    NastavHodnotu(sheetData, workbook, sharedStrings, $"ZPRAC{revIdx}", "");
                    NastavHodnotu(sheetData, workbook, sharedStrings, kontrolName, "");
                    NastavHodnotu(sheetData, workbook, sharedStrings, schvalilName, "");

                    if (!string.IsNullOrEmpty(statusCellRef))
                    {
                        SetCellValue(sheetData, Ns, statusCellRef, "", null, null);
                    }
                }
            }

            // Uložíme zpět do listu
            using (Stream stream = entry.Open())
            {
                stream.SetLength(0);
                using XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = System.Text.Encoding.UTF8 });
                doc.Save(writer);
            }
        }

        private void UpravSpotrebice(ZipArchive archiv, SharedStringsManager sharedStrings, List<Spotrebic> spotrebice)
        {
            ZipArchiveEntry entry = archiv.GetEntry("xl/worksheets/sheet2.xml") ?? throw new InvalidOperationException("V šabloně chybí soubor sheet2.xml (SEMIPROVOZ IVCHS)."); // List se spotřebiči

            XDocument doc;
            using (Stream stream = entry.Open())
            {
                doc = XDocument.Load(stream);
            }

            XElement sheetData = doc.Root.Element(Ns + "sheetData") ?? throw new InvalidOperationException("V sheet2.xml chybí element sheetData.");

            // 1. Zjistíme vzorové styly z řádku 3 (v šabloně obsahuje formátování pro prázdný řádek)
            Dictionary<string, string> vzoroveStyly = [];
            XElement vzorovyRow = sheetData.Descendants(Ns + "row").FirstOrDefault(r => r.Attribute("r")?.Value == "3");
            if (vzorovyRow != null)
            {
                foreach (XElement cell in vzorovyRow.Elements(Ns + "c"))
                {
                    string rRef = cell.Attribute("r")?.Value ?? "";
                    string colLetter = new(rRef.TakeWhile(char.IsLetter).ToArray());
                    string style = cell.Attribute("s")?.Value;
                    if (!string.IsNullOrEmpty(colLetter) && !string.IsNullOrEmpty(style))
                    {
                        vzoroveStyly[colLetter] = style;
                    }
                }
            }

            // 2. Vymažeme všechny datové řádky (řádek 4 a vyšší)
            List<XElement> radkyKeSmazani = sheetData.Descendants(Ns + "row")
                .Where(r => {
                    int rIdx = int.Parse(r.Attribute("r")?.Value ?? "0");
                    return rIdx >= 4;
                }).ToList();

            foreach (XElement row in radkyKeSmazani)
            {
                row.Remove();
            }

            // 3. Vygenerujeme nové řádky spotřebičů od řádku 4
            int aktualniRadek = 4;
            foreach (Spotrebic spotrebic in spotrebice)
            {
                var row = new XElement(Ns + "row",
                    new XAttribute("r", aktualniRadek),
                    new XAttribute("spans", "1:19")
                );

                // Naplníme buňky A až Q pro daný řádek
                PridatBunku(row, "A", aktualniRadek, spotrebic.Polozka, vzoroveStyly, sharedStrings, true); // Číslo jako string nebo číslo
                PridatBunku(row, "B", aktualniRadek, spotrebic.Rev, vzoroveStyly, sharedStrings);
                PridatBunku(row, "C", aktualniRadek, spotrebic.BalenaJednotka, vzoroveStyly, sharedStrings);
                PridatBunku(row, "D", aktualniRadek, "", vzoroveStyly, sharedStrings); // Prázdná buňka pro zachování šablony
                PridatBunku(row, "E", aktualniRadek, spotrebic.Umisteni, vzoroveStyly, sharedStrings);
                PridatBunku(row, "F", aktualniRadek, spotrebic.Tag, vzoroveStyly, sharedStrings);
                PridatBunku(row, "G", aktualniRadek, spotrebic.Zarizeni, vzoroveStyly, sharedStrings);
                PridatBunku(row, "H", aktualniRadek, spotrebic.TypVelikost, vzoroveStyly, sharedStrings);
                PridatBunku(row, "I", aktualniRadek, spotrebic.Ks, vzoroveStyly, sharedStrings);
                PridatBunku(row, "J", aktualniRadek, spotrebic.Pid, vzoroveStyly, sharedStrings);
                PridatBunku(row, "K", aktualniRadek, spotrebic.Rozvadec, vzoroveStyly, sharedStrings);
                PridatBunku(row, "L", aktualniRadek, spotrebic.Napeti, vzoroveStyly, sharedStrings, true);

                // Instalovaný příkon (Pi) - číslo
                string piVal = spotrebic.InstalovanyPi.HasValue ? spotrebic.InstalovanyPi.Value.ToString(CultureInfo.InvariantCulture) : "";
                PridatBunkuCislo(row, "M", aktualniRadek, piVal, vzoroveStyly);

                // Výpočtový příkon (Pp) - buď zadaná hodnota nebo vzorec
                if (spotrebic.VypoctovyPi.HasValue)
                {
                    string ppVal = spotrebic.VypoctovyPi.Value.ToString(CultureInfo.InvariantCulture);
                    PridatBunkuCislo(row, "N", aktualniRadek, ppVal, vzoroveStyly);
                }
                else if (spotrebic.InstalovanyPi.HasValue)
                {
                    // Generujeme vzorec MXX * 0.9
                    string vzorec = $"M{aktualniRadek}*0.9";
                    double predpokladHodnota = spotrebic.InstalovanyPi.Value * 0.9;
                    PridatBunkuVzorec(row, "N", aktualniRadek, vzorec, predpokladHodnota.ToString(CultureInfo.InvariantCulture), vzoroveStyly);
                }
                else
                {
                    PridatBunkuCislo(row, "N", aktualniRadek, "", vzoroveStyly);
                }

                PridatBunku(row, "O", aktualniRadek, spotrebic.Ivchs, vzoroveStyly, sharedStrings);
                PridatBunku(row, "P", aktualniRadek, spotrebic.StartMotoru, vzoroveStyly, sharedStrings);
                PridatBunku(row, "Q", aktualniRadek, spotrebic.Poznamka, vzoroveStyly, sharedStrings);

                sheetData.Add(row);
                aktualniRadek++;
            }

            // Uložíme zpět do listu
            using (Stream stream = entry.Open())
            {
                stream.SetLength(0);
                using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = System.Text.Encoding.UTF8 }))
                {
                    doc.Save(writer);
                }
            }
        }

        private void PridatBunku(XElement row, string colLetter, int rowIdx, string hodnota, Dictionary<string, string> vzoroveStyly, SharedStringsManager sharedStrings, bool detekovatCislo = false)
        {
            string cellRef = $"{colLetter}{rowIdx}";
            vzoroveStyly.TryGetValue(colLetter, out string styl);

            XElement cell = new XElement(Ns + "c", new XAttribute("r", cellRef));
            if (!string.IsNullOrEmpty(styl))
            {
                cell.SetAttributeValue("s", styl);
            }

            if (!string.IsNullOrEmpty(hodnota))
            {
                // Pokud chceme detekovat číslo (např. u No. nebo Napětí) a je to číslo
                if (detekovatCislo && double.TryParse(hodnota, NumberStyles.Any, CultureInfo.InvariantCulture, out double _))
                {
                    cell.Add(new XElement(Ns + "v", hodnota));
                }
                else
                {
                    int idx = sharedStrings.GetOrAdd(hodnota);
                    cell.SetAttributeValue("t", "s");
                    cell.Add(new XElement(Ns + "v", idx));
                }
            }

            row.Add(cell);
        }

        private void PridatBunkuCislo(XElement row, string colLetter, int rowIdx, string hodnota, Dictionary<string, string> vzoroveStyly)
        {
            string cellRef = $"{colLetter}{rowIdx}";
            vzoroveStyly.TryGetValue(colLetter, out string styl);

            XElement cell = new XElement(Ns + "c", new XAttribute("r", cellRef));
            if (!string.IsNullOrEmpty(styl))
            {
                cell.SetAttributeValue("s", styl);
            }

            if (!string.IsNullOrEmpty(hodnota))
            {
                cell.Add(new XElement(Ns + "v", hodnota));
            }

            row.Add(cell);
        }

        private void PridatBunkuVzorec(XElement row, string colLetter, int rowIdx, string vzorec, string cachedValue, Dictionary<string, string> vzoroveStyly)
        {
            string cellRef = $"{colLetter}{rowIdx}";
            vzoroveStyly.TryGetValue(colLetter, out string styl);

            XElement cell = new XElement(Ns + "c", new XAttribute("r", cellRef));
            if (!string.IsNullOrEmpty(styl))
            {
                cell.SetAttributeValue("s", styl);
            }

            cell.Add(new XElement(Ns + "f", vzorec));
            if (!string.IsNullOrEmpty(cachedValue))
            {
                cell.Add(new XElement(Ns + "v", cachedValue));
            }

            row.Add(cell);
        }

        private void NastavHodnotu(XElement sheetData, WorkbookManager workbook, SharedStringsManager sharedStrings, string dnName, string hodnota)
        {
            string cellRef = workbook.GetCellRef(dnName);
            if (string.IsNullOrEmpty(cellRef)) return;

            if (string.IsNullOrEmpty(hodnota))
            {
                SetCellValue(sheetData, Ns, cellRef, "", null, null);
            }
            else
            {
                // Pokud je hodnota číslo, zapíšeme ji přímo
                if (double.TryParse(hodnota, NumberStyles.Any, CultureInfo.InvariantCulture, out double _))
                {
                    SetCellValue(sheetData, Ns, cellRef, hodnota, null, null);
                }
                else
                {
                    int idx = sharedStrings.GetOrAdd(hodnota);
                    SetCellValue(sheetData, Ns, cellRef, hodnota, idx, null);
                }
            }
        }

        private static void SetCellValue(XElement sheetData, XNamespace ns, string cellRef, string val, int? stringIndex, string formula)
        {
            XElement cell = sheetData.Descendants(ns + "c").FirstOrDefault(c => c.Attribute("r")?.Value == cellRef);
            if (cell == null)
            {
                string rowStr = new string(cellRef.Where(char.IsDigit).ToArray());
                int rowIdx = int.Parse(rowStr);
                XElement row = sheetData.Descendants(ns + "row").FirstOrDefault(r => r.Attribute("r")?.Value == rowStr);
                if (row == null)
                {
                    row = new XElement(ns + "row", new XAttribute("r", rowStr));
                    XElement afterRow = sheetData.Descendants(ns + "row")
                        .FirstOrDefault(r => int.Parse(r.Attribute("r")?.Value ?? "0") > rowIdx);
                    if (afterRow != null)
                        afterRow.AddBeforeSelf(row);
                    else
                        sheetData.Add(row);
                }

                cell = new XElement(ns + "c", new XAttribute("r", cellRef));
                var cells = row.Elements(ns + "c").ToList();
                XElement afterCell = cells.FirstOrDefault(c => CompareCellRefs(c.Attribute("r")?.Value ?? "", cellRef) > 0);
                if (afterCell != null)
                    afterCell.AddBeforeSelf(cell);
                else
                    row.Add(cell);
            }

            // Vymažeme stávající vzorce a hodnoty
            cell.Element(ns + "f")?.Remove();
            cell.Element(ns + "v")?.Remove();

            if (!string.IsNullOrEmpty(formula))
            {
                cell.Add(new XElement(ns + "f", formula));
            }

            if (stringIndex.HasValue)
            {
                cell.SetAttributeValue("t", "s");
                cell.Add(new XElement(ns + "v", stringIndex.Value));
            }
            else
            {
                cell.Attribute("t")?.Remove();
                if (!string.IsNullOrEmpty(val))
                {
                    cell.Add(new XElement(ns + "v", val));
                }
            }
        }

        private static int CompareCellRefs(string a, string b)
        {
            string aCol = new(a.TakeWhile(char.IsLetter).ToArray());
            string bCol = new(b.TakeWhile(char.IsLetter).ToArray());
            if (aCol.Length != bCol.Length)
                return aCol.Length.CompareTo(bCol.Length);
            int colComp = string.CompareOrdinal(aCol, bCol);
            if (colComp != 0) return colComp;

            string aRow = new(a.SkipWhile(char.IsLetter).ToArray());
            string bRow = new(b.SkipWhile(char.IsLetter).ToArray());
            return int.Parse(aRow).CompareTo(int.Parse(bRow));
        }

        /// <summary>
        /// Vygeneruje nový Excel soubor ze šablony kabelů na základě předaných dat.
        /// </summary>
        /// <param name="sablonaCesta">Cesta k existující vzorové šabloně XLSX.</param>
        /// <param name="vystupCesta">Cesta pro uložení nově generovaného souboru.</param>
        /// <param name="coverData">Data pro vyplnění úvodního listu Cover.</param>
        /// <param name="kabely">Seznam kabelů pro uložení do listu Seznam.</param>
        public void GenerujKabel(string sablonaCesta, string vystupCesta, CoverData coverData, List<KabelPolozka> kabely)
        {
            if (!File.Exists(sablonaCesta))
            {
                throw new FileNotFoundException("Vzorová šablona nebyla nalezena.", sablonaCesta);
            }

            // Vytvoříme kopii šablony do výstupního umístění
            string adresar = Path.GetDirectoryName(vystupCesta);
            if (!string.IsNullOrEmpty(adresar) && !Directory.Exists(adresar))
            {
                Directory.CreateDirectory(adresar);
            }
            File.Copy(sablonaCesta, vystupCesta, true);

            // Otevřeme kopii pro aktualizaci XML souborů uvnitř ZIP archivu
            using ZipArchive archiv = ZipFile.Open(vystupCesta, ZipArchiveMode.Update);
            // 1. Správce sdílených řetězců (Shared Strings)
            SharedStringsManager sharedStrings = NacistSharedStrings(archiv);

            // 2. Správce sešitu (Workbook) pro vyhledání pojmenovaných oblastí
            WorkbookManager workbook = NacistWorkbook(archiv);

            // 3. Úprava listu Cover
            UpravCover(archiv, workbook, sharedStrings, coverData);

            // 4. Úprava listu s kabely
            UpravKabely(archiv, sharedStrings, kabely);

            // 5. Uložení sdílených řetězců zpět
            UlozitSharedStrings(archiv, sharedStrings);
        }

        private void UpravKabely(ZipArchive archiv, SharedStringsManager sharedStrings, List<KabelPolozka> kabely)
        {
            ZipArchiveEntry entry = archiv.GetEntry("xl/worksheets/sheet2.xml") ?? throw new InvalidOperationException("V šabloně chybí soubor sheet2.xml (Seznam).");

            XDocument doc;
            using (Stream stream = entry.Open())
            {
                doc = XDocument.Load(stream);
            }

            XElement sheetData = doc.Root.Element(Ns + "sheetData") ?? throw new InvalidOperationException("V sheet2.xml chybí element sheetData.");

            // Načteme vzorové styly z řádku 5 (který je v šabloně první datový řádek)
            Dictionary<string, string> vzoroveStyly = [];
            XElement vzorovyRow = sheetData.Descendants(Ns + "row").FirstOrDefault(r => r.Attribute("r")?.Value == "5");
            if (vzorovyRow != null)
            {
                foreach (XElement cell in vzorovyRow.Elements(Ns + "c"))
                {
                    string rRef = cell.Attribute("r")?.Value ?? "";
                    string colLetter = new(rRef.TakeWhile(char.IsLetter).ToArray());
                    string style = cell.Attribute("s")?.Value;
                    if (!string.IsNullOrEmpty(colLetter) && !string.IsNullOrEmpty(style))
                    {
                        vzoroveStyly[colLetter] = style;
                    }
                }
            }

            // Upravíme řádky 5 až 140
            for (int rIdx = 5; rIdx <= 140; rIdx++)
            {
                int kabelIndex = rIdx - 5;
                XElement row = sheetData.Descendants(Ns + "row").FirstOrDefault(r => r.Attribute("r")?.Value == rIdx.ToString());
                if (row == null)
                {
                    row = new XElement(Ns + "row",
                        new XAttribute("r", rIdx),
                        new XAttribute("spans", "1:16")
                    );
                    XElement prevRow = sheetData.Descendants(Ns + "row")
                        .FirstOrDefault(r => int.Parse(r.Attribute("r")?.Value ?? "0") == rIdx - 1);
                    if (prevRow != null)
                        prevRow.AddAfterSelf(row);
                    else
                        sheetData.Add(row);
                }

                if (kabelIndex < kabely.Count)
                {
                    KabelPolozka kabel = kabely[kabelIndex];
                    // Zápis hodnot do buněk A až K
                    NastavBunku(row, "A", rIdx, kabel.Polozka, vzoroveStyly, sharedStrings, detekovatCislo: true);
                    NastavBunku(row, "B", rIdx, kabel.OznaceniKabeluPuvodni, vzoroveStyly, sharedStrings);
                    NastavBunku(row, "C", rIdx, kabel.CisloKabelu, vzoroveStyly, sharedStrings);
                    NastavBunku(row, "D", rIdx, kabel.KabelTyp, vzoroveStyly, sharedStrings);
                    NastavBunku(row, "E", rIdx, kabel.Prurez, vzoroveStyly, sharedStrings);

                    string delkaVal = kabel.Delka.HasValue ? kabel.Delka.Value.ToString(CultureInfo.InvariantCulture) : "";
                    NastavBunkuCislo(row, "F", rIdx, delkaVal, vzoroveStyly);

                    NastavBunku(row, "G", rIdx, kabel.ZeZarizeni, vzoroveStyly, sharedStrings);
                    NastavBunku(row, "H", rIdx, kabel.UkonceniZe, vzoroveStyly, sharedStrings);
                    NastavBunku(row, "I", rIdx, kabel.DoZarizeni, vzoroveStyly, sharedStrings);
                    NastavBunku(row, "J", rIdx, kabel.UkonceniDo, vzoroveStyly, sharedStrings);
                    NastavBunku(row, "K", rIdx, kabel.Poznamka, vzoroveStyly, sharedStrings);
                }
                else
                {
                    // Vyčistíme buňky A až K (ale zachováme styly a strukturu řádku)
                    VyčistitBunky(row, rIdx, ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]);
                }
            }

            // Uložíme zpět do listu
            using (Stream stream = entry.Open())
            {
                stream.SetLength(0);
                using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = System.Text.Encoding.UTF8 }))
                {
                    doc.Save(writer);
                }
            }
        }

        private void NastavBunku(XElement row, string colLetter, int rowIdx, string hodnota, Dictionary<string, string> vzoroveStyly, SharedStringsManager sharedStrings, bool detekovatCislo = false)
        {
            string cellRef = $"{colLetter}{rowIdx}";
            XElement cell = row.Elements(Ns + "c").FirstOrDefault(c => c.Attribute("r")?.Value == cellRef);
            if (cell == null)
            {
                cell = new XElement(Ns + "c", new XAttribute("r", cellRef));
                vzoroveStyly.TryGetValue(colLetter, out string styl);
                if (!string.IsNullOrEmpty(styl))
                {
                    cell.SetAttributeValue("s", styl);
                }
                row.Add(cell);
            }

            cell.Element(Ns + "f")?.Remove();
            cell.Element(Ns + "v")?.Remove();

            if (!string.IsNullOrEmpty(hodnota))
            {
                if (detekovatCislo && double.TryParse(hodnota, NumberStyles.Any, CultureInfo.InvariantCulture, out double _))
                {
                    cell.Attribute("t")?.Remove();
                    cell.Add(new XElement(Ns + "v", hodnota));
                }
                else
                {
                    int idx = sharedStrings.GetOrAdd(hodnota);
                    cell.SetAttributeValue("t", "s");
                    cell.Add(new XElement(Ns + "v", idx));
                }
            }
            else
            {
                cell.Attribute("t")?.Remove();
            }
        }

        private void NastavBunkuCislo(XElement row, string colLetter, int rowIdx, string hodnota, Dictionary<string, string> vzoroveStyly)
        {
            string cellRef = $"{colLetter}{rowIdx}";
            XElement cell = row.Elements(Ns + "c").FirstOrDefault(c => c.Attribute("r")?.Value == cellRef);
            if (cell == null)
            {
                cell = new XElement(Ns + "c", new XAttribute("r", cellRef));
                vzoroveStyly.TryGetValue(colLetter, out string styl);
                if (!string.IsNullOrEmpty(styl))
                {
                    cell.SetAttributeValue("s", styl);
                }
                row.Add(cell);
            }

            cell.Element(Ns + "f")?.Remove();
            cell.Element(Ns + "v")?.Remove();

            if (!string.IsNullOrEmpty(hodnota))
            {
                cell.Attribute("t")?.Remove();
                cell.Add(new XElement(Ns + "v", hodnota));
            }
            else
            {
                cell.Attribute("t")?.Remove();
            }
        }

        private void VyčistitBunky(XElement row, int rowIdx, string[] colLetters)
        {
            foreach (string colLetter in colLetters)
            {
                string cellRef = $"{colLetter}{rowIdx}";
                XElement cell = row.Elements(Ns + "c").FirstOrDefault(c => c.Attribute("r")?.Value == cellRef);
                if (cell != null)
                {
                    cell.Element(Ns + "f")?.Remove();
                    cell.Element(Ns + "v")?.Remove();
                    cell.Attribute("t")?.Remove();
                }
            }
        }
    }

    /// <summary>
    /// Správce pojmenovaných oblastí v sešitu.
    /// </summary>
    public class WorkbookManager
    {
        private readonly Dictionary<string, string> definedNames = new(StringComparer.OrdinalIgnoreCase);

        public WorkbookManager(Stream stream)
        {
            XDocument doc = XDocument.Load(stream);
            XNamespace ns = doc.Root.GetDefaultNamespace();
            foreach (XElement dn in doc.Descendants(ns + "definedName"))
            {
                string name = dn.Attribute("name")?.Value ?? "";
                string val = dn.Value;
                if (!string.IsNullOrEmpty(name) && val.Contains("!") && !val.StartsWith("#"))
                {
                    string[] parts = val.Split('!');
                    if (parts.Length > 1 && !string.IsNullOrEmpty(parts[1]))
                    {
                        string cellRef = parts[1].Replace("$", "");
                        definedNames[name] = cellRef;
                    }
                }
            }
        }

        public string GetCellRef(string name)
        {
            definedNames.TryGetValue(name, out string cellRef);
            return cellRef;
        }
    }

    /// <summary>
    /// Správce sdílených řetězců (Shared Strings).
    /// </summary>
    public class SharedStringsManager
    {
        private readonly XNamespace ns;
        private readonly List<string> strings = [];

        public SharedStringsManager(Stream stream)
        {
            XDocument doc = XDocument.Load(stream);
            ns = doc.Root.GetDefaultNamespace();
            foreach (XElement si in doc.Descendants(ns + "si"))
            {
                XElement tNode = si.Element(ns + "t");
                if (tNode != null)
                {
                    strings.Add(tNode.Value);
                }
                else
                {
                    string val = string.Concat(si.Descendants(ns + "t").Select(t => t.Value));
                    strings.Add(val);
                }
            }
        }

        public int GetOrAdd(string text)
        {
            if (text == null) text = string.Empty;
            int idx = strings.IndexOf(text);
            if (idx >= 0) return idx;
            strings.Add(text);
            return strings.Count - 1;
        }

        public void Save(Stream stream)
        {
            stream.SetLength(0);
            XElement sst = new XElement(ns + "sst",
                new XAttribute("count", strings.Count),
                new XAttribute("uniqueCount", strings.Count),
                strings.Select(s => new XElement(ns + "si", new XElement(ns + "t", s)))
            );
            XDocument newDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), sst);
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = System.Text.Encoding.UTF8 }))
            {
                newDoc.Save(writer);
            }
        }
    }
}
