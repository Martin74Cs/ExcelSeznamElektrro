using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Knihovna.KabelyXls
{
    /// <summary>
    /// Generátor Excelu, který provádí úpravy v šabloně pomocí knihovny ClosedXML.
    /// Zjednodušená verze nahrazující přímou manipulaci s OpenXML balíčkem.
    /// </summary>
    public class ExcelGenerator
    {
        /// <summary>
        /// Vygeneruje nový Excel soubor ze šablony na základě předaných dat pro spotřebiče.
        /// </summary>
        /// <param name="sablonaCesta">Cesta k existující vzorové šabloně XLSX.</param>
        /// <param name="vystupCesta">Cesta pro uložení nově generovaného souboru.</param>
        /// <param name="coverData">Data pro vyplnění úvodního listu Cover.</param>
        /// <param name="spotrebice">Seznam spotřebičů pro uložení do listu.</param>
        public void Generuj(string sablonaCesta, string vystupCesta, CoverData coverData, List<Spotrebic> spotrebice)
        {
            if (!File.Exists(sablonaCesta))
            {
                throw new FileNotFoundException("Vzorová šablona nebyla nalezena.", sablonaCesta);
            }

            // Vytvoříme kopii šablony do výstupního umístění
            string? adresar = Path.GetDirectoryName(vystupCesta);
            if (!string.IsNullOrEmpty(adresar) && !Directory.Exists(adresar))
            {
                Directory.CreateDirectory(adresar);
            }
            File.Copy(sablonaCesta, vystupCesta, true);

            // Otevřeme a upravíme pomocí ClosedXML
            using (XLWorkbook workbook = new(vystupCesta))
            {
                // 1. Úprava listu Cover
                UpravCover(workbook, coverData);

                // 2. Úprava listu se spotřebiči
                UpravSpotrebice(workbook, spotrebice);

                // Uložení změn
                workbook.Save();
            }
        }

        /// <summary>
        /// Vygeneruje nový Excel soubor ze šablony na základě předaných dat pro kabely.
        /// </summary>
        /// <param name="sablonaCesta">Cesta k existující vzorové šabloně XLSX.</param>
        /// <param name="vystupCesta">Cesta pro uložení nově generovaného souboru.</param>
        /// <param name="coverData">Data pro vyplnění úvodního listu Cover.</param>
        /// <param name="kabely">Seznam kabelů pro uložení do listu.</param>
        public void GenerujKabel(string sablonaCesta, string vystupCesta, CoverData coverData, List<KabelPolozka> kabely)
        {
            if (!File.Exists(sablonaCesta))
            {
                throw new FileNotFoundException("Vzorová šablona nebyla nalezena.", sablonaCesta);
            }

            // Vytvoříme kopii šablony do výstupního umístění
            string? adresar = Path.GetDirectoryName(vystupCesta);
            if (!string.IsNullOrEmpty(adresar) && !Directory.Exists(adresar))
            {
                Directory.CreateDirectory(adresar);
            }
            File.Copy(sablonaCesta, vystupCesta, true);

            // Otevřeme a upravíme pomocí ClosedXML
            using (XLWorkbook workbook = new(vystupCesta))
            {
                // 1. Úprava listu Cover
                UpravCover(workbook, coverData);

                // 2. Úprava listu s kabely
                UpravKabely(workbook, kabely);

                // Uložení změn
                workbook.Save();
            }
        }

        /// <summary>
        /// Vyplní úvodní list Cover podle zadaných dat.
        /// </summary>
        private void UpravCover(XLWorkbook workbook, CoverData coverData)
        {
            IXLWorksheet ws = workbook.Worksheet(1); // Cover list je první

            void NastavHodnotu(string name, string hodnota)
            {
                IXLCell? cell = GetCellByDefinedName(workbook, name);
                if (cell != null)
                {
                    if (string.IsNullOrEmpty(hodnota))
                    {
                        cell.Value = string.Empty;
                    }
                    else if (double.TryParse(hodnota, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                    {
                        cell.Value = num;
                    }
                    else
                    {
                        cell.Value = hodnota;
                    }
                }
            }

            // Vyplnění základních definovaných názvů
            NastavHodnotu("ZAKAZNIK", coverData.Zakaznik);
            NastavHodnotu("PROJEKT", coverData.Projekt);
            NastavHodnotu("NAZEV", coverData.Nazev);
            NastavHodnotu("_NA4", coverData.DokumentNazev);
            NastavHodnotu("_CAS4", coverData.Technologie);
            NastavHodnotu("CDOK1", coverData.CistyDokumentTyp);
            NastavHodnotu("CDOK2", coverData.CistyDokumentCislo);
            NastavHodnotu("REV", coverData.Revize);

            // Vyplnění revizní tabulky (max 6 řádků)
            for (int i = 0; i < 6; i++)
            {
                string revIdx = (i + 1).ToString();
                string revName = $"_REV{revIdx}";
                string datName = $"_DAT{revIdx}";
                string popName = $"_POP{revIdx}";
                string zpracName = $"ZPRAC{revIdx}";
                string kontrolName = $"KONTROL{revIdx}";
                string schvalilName = $"SCHVALIL{revIdx}";

                IXLCell? revCell = GetCellByDefinedName(workbook, revName);

                if (i < coverData.RevizeSeznam.Count)
                {
                    RevizeInfo rev = coverData.RevizeSeznam[i];
                    NastavHodnotu(revName, rev.Revize);
                    NastavHodnotu(datName, rev.Datum);
                    NastavHodnotu(popName, rev.PopisRevize);
                    NastavHodnotu(zpracName, rev.Zpacoval);
                    NastavHodnotu(kontrolName, rev.Kontroloval);
                    NastavHodnotu(schvalilName, rev.Schvalil);

                    // Zapíšeme status revize do sloupce D ve stejném řádku jako _REVx
                    if (revCell != null)
                    {
                        ws.Cell(revCell.Address.RowNumber, "D").Value = rev.Stat;
                    }
                }
                else
                {
                    // Vyčistíme zbylé řádky
                    NastavHodnotu(revName, string.Empty);
                    NastavHodnotu(datName, string.Empty);
                    NastavHodnotu(popName, string.Empty);
                    NastavHodnotu(zpracName, string.Empty);
                    NastavHodnotu(kontrolName, string.Empty);
                    NastavHodnotu(schvalilName, string.Empty);

                    if (revCell != null)
                    {
                        ws.Cell(revCell.Address.RowNumber, "D").Value = string.Empty;
                    }
                }
            }
        }

        /// <summary>
        /// Vyplní seznam spotřebičů v druhém listu šablony.
        /// </summary>
        private void UpravSpotrebice(XLWorkbook workbook, List<Spotrebic> spotrebice)
        {
            IXLWorksheet ws = workbook.Worksheet(2); // List se spotřebiči

            // Zjistíme vzorový řádek (řádek 3 v šabloně) pro zkopírování stylů
            IXLRow templateRow = ws.Row(3);

            // Vymažeme všechny datové řádky od řádku 4 dále
            int posledniRadek = ws.LastRowUsed()?.RowNumber() ?? 4;
            if (posledniRadek >= 4)
            {
                ws.Rows(4, posledniRadek).Delete();
            }

            // Vygenerujeme nové řádky spotřebičů od řádku 4
            int aktualniRadek = 4;
            foreach (Spotrebic spotrebic in spotrebice)
            {
                IXLRow row = ws.Row(aktualniRadek);

                // Zkopírujeme styly buněk A až M (sloupce 1 až 13)
                for (int col = 1; col <= 13; col++)
                {
                    row.Cell(col).Style = templateRow.Cell(col).Style;
                }

                // Naplnění hodnot podle požadovaného pořadí sloupců
                NastavBunku(row.Cell("A"), spotrebic.Polozka, true);
                NastavBunku(row.Cell("B"), spotrebic.Rev);
                NastavBunku(row.Cell("C"), spotrebic.BalenaJednotka);
                NastavBunku(row.Cell("D"), spotrebic.Umisteni);
                NastavBunku(row.Cell("E"), spotrebic.Tag);
                NastavBunku(row.Cell("F"), spotrebic.Popis);
                NastavBunku(row.Cell("G"), spotrebic.TypVelikost);
                NastavBunku(row.Cell("H"), spotrebic.Rozvadec);
                NastavBunku(row.Cell("I"), spotrebic.Napeti, true);
                NastavBunku(row.Cell("J"), spotrebic.InstalovanyPi, true);
                // K: Příkon Pp (Výpočtový příkon)
                if (double.TryParse(spotrebic.InstalovanyPi, out double InstalovanyPi))
                {
                    row.Cell("K").FormulaA1 = $"J{aktualniRadek}*0.9";
                }
                else
                {
                    row.Cell("K").Value = string.Empty;
                }

                NastavBunku(row.Cell("L"), spotrebic.StartMotoru);
                NastavBunku(row.Cell("M"), spotrebic.Poznamka);

                aktualniRadek++;
            }
        }

        /// <summary>
        /// Vyplní seznam kabelů v druhém listu šablony.
        /// </summary>
        private void UpravKabely(XLWorkbook workbook, List<KabelPolozka> kabely)
        {
            IXLWorksheet ws = workbook.Worksheet(2); // List s kabely

            // Načteme vzorový řádek (řádek 5 v šabloně)
            IXLRow templateRow = ws.Row(5);

            // Upravíme řádky od indexu 5 dále (minimálně do 140, případně více dle počtu kabelů)
            int maxRadek = Math.Max(140, 4 + kabely.Count);

            for (int rIdx = 5; rIdx <= maxRadek; rIdx++)
            {
                int kabelIndex = rIdx - 5;
                IXLRow row = ws.Row(rIdx);

                if (kabelIndex < kabely.Count)
                {
                    KabelPolozka kabel = kabely[kabelIndex];

                    // Zkopírujeme styly buněk A až K
                    for (int col = 1; col <= 11; col++)
                    {
                        row.Cell(col).Style = templateRow.Cell(col).Style;
                    }

                    // Nastavíme hodnoty
                    NastavBunku(row.Cell("A"), kabel.Polozka, true);
                    NastavBunku(row.Cell("B"), kabel.Revize);
                    NastavBunku(row.Cell("C"), kabel.CisloKabelu);
                    NastavBunku(row.Cell("D"), kabel.KabelTyp);
                    NastavBunku(row.Cell("E"), kabel.Prurez);
                    NastavBunku(row.Cell("F"), kabel.Delka, true);
                    NastavBunku(row.Cell("G"), kabel.ZeZarizeni);
                    NastavBunku(row.Cell("H"), kabel.UkonceniZe);
                    NastavBunku(row.Cell("I"), kabel.DoZarizeni);
                    NastavBunku(row.Cell("J"), kabel.UkonceniDo);
                    NastavBunku(row.Cell("K"), kabel.Poznamka);
                }
                else
                {
                    // Vyčistíme buňky A až K (ale zachováme styly)
                    for (int col = 1; col <= 11; col++)
                    {
                        row.Cell(col).Value = string.Empty;
                    }
                }
            }
        }

        /// <summary>
        /// Pomocná metoda pro bezpečné nastavení hodnoty buňky s volitelnou detekcí čísla.
        /// </summary>
        private static void NastavBunku(IXLCell cell, string hodnota, bool detekovatCislo = false)
        {
            if (string.IsNullOrEmpty(hodnota))
            {
                cell.Value = string.Empty;
            }
            else if (detekovatCislo && double.TryParse(hodnota, NumberStyles.Any, CultureInfo.InvariantCulture, out double cislo))
            {
                cell.Value = cislo;
            }
            else
            {
                cell.Value = hodnota;
            }
        }

        /// <summary>
        /// Pomocná metoda pro nalezení buňky podle definovaného názvu v sešitu.
        /// </summary>
        private static IXLCell? GetCellByDefinedName(XLWorkbook workbook, string name)
        {
            IXLDefinedName? definedName = workbook.DefinedNames.FirstOrDefault(dn => string.Equals(dn.Name, name, StringComparison.OrdinalIgnoreCase));
            return definedName?.Ranges.FirstOrDefault()?.FirstCell();
        }
    }
}
