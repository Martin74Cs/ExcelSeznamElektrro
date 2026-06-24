using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;
using MigraDocDoc = MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace Knihovna.Export
{
    /// <summary>
    /// Reprezentuje pojmenovanou sekci s daty pro sloučený export.
    /// </summary>
    /// <typeparam name="T">Typ objektů v sekci.</typeparam>
    public class ExportSection<T>
    {
        /// <summary>
        /// Nadpis sekce (zobrazí se nad příslušnou tabulkou).
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// Seznam datových položek pro tuto sekci.
        /// </summary>
        public List<T> Data { get; set; } = [];
    }

    /// <summary>
    /// Třída zajišťující exporty běžných i sloučených seznamů do různých formátů (XLSX, CSV, XML, HTML, PDF, DOCX).
    /// </summary>
    public static class SloucenyExporter
    {
        #region Pomocné metody pro reflexi a formátování

        /// <summary>
        /// Získá seznam vlastností typu T, které mají být exportovány.
        /// Pokud jsou specifikovány sloupce, filtruje a seřadí je podle nich.
        /// </summary>
        private static List<PropertyInfo> GetPropertiesToExport<T>(string[]? columns)
        {
            List<PropertyInfo> allProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                                         .Where(p => p.GetIndexParameters().Length == 0)
                                                         .ToList();

            if (columns != null && columns.Length > 0)
            {
                return columns
                    .Select(colName => allProperties.FirstOrDefault(p => string.Equals(p.Name, colName, StringComparison.OrdinalIgnoreCase)))
                    .Where(p => p != null)
                    .Cast<PropertyInfo>()
                    .ToList();
            }

            return allProperties;
        }

        /// <summary>
        /// Vrátí text záhlaví sloupce na základě DisplayAttribute (pokud existuje) nebo názvu vlastnosti.
        /// </summary>
        private static string GetHeaderName(PropertyInfo prop)
        {
            DisplayAttribute? displayAttr = prop.GetCustomAttribute<DisplayAttribute>();
            string name = displayAttr?.Name ?? prop.Name;
            string? unit = displayAttr?.Prompt; // Promt se v projektu používá jako jednotka
            
            return string.IsNullOrEmpty(unit) ? name : $"{name} [{unit}]";
        }

        /// <summary>
        /// Zformátuje hodnotu vlastnosti pro textové výstupy.
        /// </summary>
        private static string FormatValue(object? value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            return value switch
            {
                DateTime dt => dt.ToString("dd.MM.yyyy HH:mm:ss"),
                bool b => b ? "Ano" : "Ne",
                _ => value.ToString() ?? string.Empty
            };
        }

        #endregion

        #region Běžný export jednoho seznamu do XLSX (ClosedXML)

        /// <summary>
        /// Exportuje jeden seznam dat do Excelu (.xlsx) pomocí ClosedXML.
        /// </summary>
        /// <typeparam name="T">Typ exportovaných dat.</typeparam>
        /// <param name="data">Seznam datových položek.</param>
        /// <param name="xlsxPath">Cesta, kam se má soubor uložit.</param>
        /// <param name="title">Volitelný nadpis tabulky.</param>
        /// <param name="columns">Volitelný seznam sloupců k exportu.</param>
        public static void SaveXlsxGen<T>(this List<T> data, string xlsxPath, string? title = null, string[]? columns = null) where T : new()
        {
            if (data == null || data.Count == 0)
            {
                return;
            }

            if (!Soubory.CanSaveFile(xlsxPath))
            {
                return;
            }

            using XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet ws = workbook.Worksheets.Add("Seznam");

            List<PropertyInfo> properties = GetPropertiesToExport<T>(columns);
            int currentRow = 1;

            // Zápis hlavního nadpisu, pokud je zadán
            if (!string.IsNullOrWhiteSpace(title))
            {
                IXLCell cellTitle = ws.Cell(currentRow, 1);
                cellTitle.Value = title;
                cellTitle.Style.Font.Bold = true;
                cellTitle.Style.Font.FontSize = 16;
                currentRow += 2;
            }

            // Zápis hlavičky tabulky
            for (int col = 0; col < properties.Count; col++)
            {
                IXLCell cell = ws.Cell(currentRow, col + 1);
                cell.Value = GetHeaderName(properties[col]);
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#2B6CB0");
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            currentRow++;

            // Zápis dat
            for (int row = 0; row < data.Count; row++)
            {
                T item = data[row];
                bool isEven = row % 2 == 1;

                for (int col = 0; col < properties.Count; col++)
                {
                    IXLCell cell = ws.Cell(currentRow, col + 1);
                    object? rawVal = properties[col].GetValue(item);

                    if (rawVal != null)
                    {
                        string strVal = rawVal.ToString() ?? string.Empty;
                        
                        // Detekce vzorců (začíná na "=")
                        if (strVal.StartsWith("=") && strVal.Length > 1)
                        {
                            cell.FormulaA1 = strVal.Substring(1);
                        }
                        // Detekce čísel
                        else if (double.TryParse(strVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double numVal))
                        {
                            cell.Value = numVal;
                        }
                        else
                        {
                            cell.Value = strVal;
                        }
                    }

                    // Styl řádku (střídavé řádky a jemné ohraničení)
                    if (isEven)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F7FAFC");
                    }
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#E2E8F0");
                }
                currentRow++;
            }

            // Automatická šířka sloupců
            ws.Columns().AdjustToContents();

            // Uložení sešitu
            workbook.SaveAs(xlsxPath);
            Console.WriteLine($"Hotovo! Soubor XLSX byl uložen do {Path.GetFileName(xlsxPath)}");
        }

        #endregion

        #region Sloučený export do XLSX (ClosedXML)

        /// <summary>
        /// Exportuje více seznamů (sekcí) do jednoho listu Excelu (.xlsx) pod sebe s nadpisy a formátováním.
        /// </summary>
        public static void SaveXlsxSections<T>(string xlsxPath, List<ExportSection<T>> sections, string? mainTitle = null, string[]? columns = null) where T : new()
        {
            if (sections == null || sections.Count == 0)
            {
                return;
            }

            if (!Soubory.CanSaveFile(xlsxPath))
            {
                return;
            }

            using XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet ws = workbook.Worksheets.Add("Sloučený seznam");

            List<PropertyInfo> properties = GetPropertiesToExport<T>(columns);
            int currentRow = 1;

            // Zápis hlavního nadpisu dokumentu
            if (!string.IsNullOrWhiteSpace(mainTitle))
            {
                IXLCell cellTitle = ws.Cell(currentRow, 1);
                cellTitle.Value = mainTitle;
                cellTitle.Style.Font.Bold = true;
                cellTitle.Style.Font.FontSize = 18;
                currentRow += 2;
            }

            foreach (ExportSection<T> section in sections)
            {
                // Nadpis sekce
                IXLCell cellSection = ws.Cell(currentRow, 1);
                cellSection.Value = section.Title;
                cellSection.Style.Font.Bold = true;
                cellSection.Style.Font.FontSize = 14;
                cellSection.Style.Font.FontColor = XLColor.FromHtml("#2C5282");
                currentRow++;

                // Hlavička sekce
                for (int col = 0; col < properties.Count; col++)
                {
                    IXLCell cell = ws.Cell(currentRow, col + 1);
                    cell.Value = GetHeaderName(properties[col]);
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#2B6CB0");
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                currentRow++;

                // Datové řádky sekce
                if (section.Data == null || section.Data.Count == 0)
                {
                    // Pokud je sekce prázdná
                    IXLCell emptyCell = ws.Cell(currentRow, 1);
                    emptyCell.Value = "Žádná data v této sekci.";
                    emptyCell.Style.Font.Italic = true;
                    currentRow++;
                }
                else
                {
                    for (int row = 0; row < section.Data.Count; row++)
                    {
                        T item = section.Data[row];
                        bool isEven = row % 2 == 1;

                        for (int col = 0; col < properties.Count; col++)
                        {
                            IXLCell cell = ws.Cell(currentRow, col + 1);
                            object? rawVal = properties[col].GetValue(item);

                            if (rawVal != null)
                            {
                                string strVal = rawVal.ToString() ?? string.Empty;
                                if (strVal.StartsWith("=") && strVal.Length > 1)
                                {
                                    cell.FormulaA1 = strVal.Substring(1);
                                }
                                else if (double.TryParse(strVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double numVal))
                                {
                                    cell.Value = numVal;
                                }
                                else
                                {
                                    cell.Value = strVal;
                                }
                            }

                            if (isEven)
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F7FAFC");
                            }
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#E2E8F0");
                        }
                        currentRow++;
                    }
                }

                // Mezera za sekcí
                currentRow += 2;
            }

            // Přizpůsobení šířky sloupců
            ws.Columns().AdjustToContents();

            workbook.SaveAs(xlsxPath);
            Console.WriteLine($"Hotovo! Sloučený soubor XLSX byl uložen do {Path.GetFileName(xlsxPath)}");
        }

        #endregion

        #region Sloučený export do CSV

        /// <summary>
        /// Exportuje sloučený seznam do jednoho CSV souboru s oddělovači sekcí.
        /// </summary>
        public static void SaveCsvSections<T>(string csvPath, List<ExportSection<T>> sections, string[]? columns = null) where T : new()
        {
            if (sections == null || sections.Count == 0)
            {
                return;
            }

            if (!Soubory.CanSaveFile(csvPath))
            {
                return;
            }

            List<PropertyInfo> properties = GetPropertiesToExport<T>(columns);

            using StreamWriter writer = new StreamWriter(csvPath, false, new UTF8Encoding(true));

            foreach (ExportSection<T> section in sections)
            {
                // Nadpis sekce
                writer.WriteLine($"# === {section.Title.ToUpper()} ===");
                
                // Hlavička sekce
                writer.WriteLine(string.Join(";", properties.Select(p => p.Name)));

                // Data sekce
                if (section.Data != null)
                {
                    foreach (T item in section.Data)
                    {
                        var values = properties.Select(p =>
                        {
                            object? valObj = p.GetValue(item);
                            string value = FormatValue(valObj)
                                .Replace("\"", "\"\"")
                                .Replace("\n", " ")
                                .Replace("\r", " ");
                            return $"\"{value}\"";
                        });
                        writer.WriteLine(string.Join(";", values));
                    }
                }

                // Prázdný řádek pro oddělení sekcí
                writer.WriteLine();
            }

            Console.WriteLine($"Hotovo! Sloučený soubor CSV byl uložen do {Path.GetFileName(csvPath)}");
        }

        #endregion

        #region Sloučený export do XML

        /// <summary>
        /// Exportuje sloučený seznam do XML souboru se strukturovanými sekcemi.
        /// </summary>
        public static void SaveXmlSections<T>(string xmlPath, List<ExportSection<T>> sections, string[]? columns = null) where T : new()
        {
            if (sections == null || sections.Count == 0)
            {
                return;
            }

            if (!Soubory.CanSaveFile(xmlPath))
            {
                return;
            }

            List<PropertyInfo> properties = GetPropertiesToExport<T>(columns);
            string rootName = $"SloucenyExportOf{typeof(T).Name}";
            string itemName = typeof(T).Name;

            XDocument xDoc = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement(rootName,
                    sections.Select(section =>
                        new XElement("Sekce",
                            new XAttribute("Nazev", section.Title),
                            section.Data.Select(item =>
                                new XElement(itemName,
                                    properties.Select(p =>
                                    {
                                        object? val = p.GetValue(item);
                                        return new XElement(p.Name, val ?? string.Empty);
                                    })
                                )
                            )
                        )
                    )
                )
            );

            xDoc.Save(xmlPath);
            Console.WriteLine($"Hotovo! Sloučený soubor XML byl uložen do {Path.GetFileName(xmlPath)}");
        }

        #endregion

        #region Sloučený export do HTML

        /// <summary>
        /// Exportuje sloučený seznam do jednoho HTML souboru s tabulkami pro každou sekci a CSS stylem.
        /// </summary>
        public static void SaveHtmlSections<T>(string htmlPath, List<ExportSection<T>> sections, string? mainTitle = null, string[]? columns = null) where T : new()
        {
            if (sections == null || sections.Count == 0)
            {
                return;
            }

            if (!Soubory.CanSaveFile(htmlPath))
            {
                return;
            }

            List<PropertyInfo> properties = GetPropertiesToExport<T>(columns);
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html><head><meta charset=\"UTF-8\"><title>Sloučený seznam</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; background-color: #f9f9f9; color: #333; }");
            sb.AppendLine("h1 { color: #2B6CB0; border-bottom: 2px solid #2B6CB0; padding-bottom: 10px; }");
            sb.AppendLine("h2 { color: #2C5282; margin-top: 30px; margin-bottom: 10px; }");
            sb.AppendLine("table { border-collapse: collapse; width: 100%; background-color: #fff; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }");
            sb.AppendLine("th, td { border: 1px solid #E2E8F0; padding: 10px 12px; text-align: left; }");
            sb.AppendLine("th { background-color: #2B6CB0; color: white; font-weight: bold; }");
            sb.AppendLine("tr:nth-child(even) { background-color: #F7FAFC; }");
            sb.AppendLine(".empty { font-style: italic; color: #718096; padding: 10px; }");
            sb.AppendLine("</style></head><body>");

            if (!string.IsNullOrEmpty(mainTitle))
            {
                sb.AppendLine($"<h1>{mainTitle}</h1>");
            }

            foreach (ExportSection<T> section in sections)
            {
                sb.AppendLine($"<h2>{section.Title}</h2>");

                if (section.Data == null || section.Data.Count == 0)
                {
                    sb.AppendLine("<div class=\"empty\">Tato sekce neobsahuje žádná data.</div>");
                    continue;
                }

                sb.AppendLine("<table><thead><tr>");
                // Záhlaví tabulky
                foreach (PropertyInfo prop in properties)
                {
                    sb.AppendLine($"<th>{System.Net.WebUtility.HtmlEncode(GetHeaderName(prop))}</th>");
                }
                sb.AppendLine("</tr></thead><tbody>");

                // Data tabulky
                foreach (T item in section.Data)
                {
                    sb.AppendLine("<tr>");
                    foreach (PropertyInfo prop in properties)
                    {
                        object? valObj = prop.GetValue(item);
                        string valStr = FormatValue(valObj);
                        sb.AppendLine($"<td>{System.Net.WebUtility.HtmlEncode(valStr)}</td>");
                    }
                    sb.AppendLine("</tr>");
                }
                sb.AppendLine("</tbody></table>");
            }

            sb.AppendLine("</body></html>");

            File.WriteAllText(htmlPath, sb.ToString(), Encoding.UTF8);
            Console.WriteLine($"Hotovo! Sloučený soubor HTML byl uložen do {Path.GetFileName(htmlPath)}");
        }

        #endregion

        #region Sloučený export do PDF (MigraDoc)

        /// <summary>
        /// Exportuje sloučený seznam do PDF pomocí knihovny MigraDoc.
        /// </summary>
        public static void SavePdfSections<T>(string pdfPath, List<ExportSection<T>> sections, string? mainTitle = null, string[]? columns = null) where T : new()
        {
            if (sections == null || sections.Count == 0)
            {
                return;
            }

            if (!Soubory.CanSaveFile(pdfPath))
            {
                return;
            }

            MigraDocDoc.Document document = new MigraDocDoc.Document();
            MigraDocDoc.Section pdfSection = document.AddSection();

            // Nastavení A4 Landscape
            pdfSection.PageSetup.PageFormat = MigraDocDoc.PageFormat.A4;
            pdfSection.PageSetup.Orientation = MigraDocDoc.Orientation.Landscape;
            pdfSection.PageSetup.TopMargin = MigraDocDoc.Unit.FromCentimeter(1.5);
            pdfSection.PageSetup.BottomMargin = MigraDocDoc.Unit.FromCentimeter(1.5);
            pdfSection.PageSetup.LeftMargin = MigraDocDoc.Unit.FromCentimeter(1.5);
            pdfSection.PageSetup.RightMargin = MigraDocDoc.Unit.FromCentimeter(1.5);

            // Nastavení základního stylu písma
            MigraDocDoc.Style normalStyle = document.Styles["Normal"];
            normalStyle.Font.Name = "Calibri";
            normalStyle.Font.Size = 9;

            // Hlavní nadpis dokumentu
            if (!string.IsNullOrWhiteSpace(mainTitle))
            {
                MigraDocDoc.Paragraph titleParagraph = pdfSection.AddParagraph(mainTitle);
                titleParagraph.Format.Font.Size = 18;
                titleParagraph.Format.Font.Bold = true;
                titleParagraph.Format.Font.Color = MigraDocDoc.Colors.DarkBlue;
                titleParagraph.Format.SpaceAfter = MigraDocDoc.Unit.FromCentimeter(0.5);
            }

            List<PropertyInfo> properties = GetPropertiesToExport<T>(columns);

            foreach (ExportSection<T> dataSection in sections)
            {
                // Nadpis sekce
                MigraDocDoc.Paragraph secTitleParagraph = pdfSection.AddParagraph(dataSection.Title);
                secTitleParagraph.Format.Font.Size = 12;
                secTitleParagraph.Format.Font.Bold = true;
                secTitleParagraph.Format.SpaceBefore = MigraDocDoc.Unit.FromCentimeter(0.5);
                secTitleParagraph.Format.SpaceAfter = MigraDocDoc.Unit.FromCentimeter(0.2);
                secTitleParagraph.Format.KeepWithNext = true;

                if (dataSection.Data == null || dataSection.Data.Count == 0)
                {
                    MigraDocDoc.Paragraph emptyParagraph = pdfSection.AddParagraph("Žádná data v této sekci.");
                    emptyParagraph.Format.Font.Italic = true;
                    emptyParagraph.Format.SpaceAfter = MigraDocDoc.Unit.FromCentimeter(0.4);
                    continue;
                }

                // Výpočet šířek sloupců (podobně jako v pdf.cs)
                double[] maxLen = new double[properties.Count];
                for (int i = 0; i < properties.Count; i++)
                {
                    maxLen[i] = GetHeaderName(properties[i]).Length;
                }

                foreach (T item in dataSection.Data)
                {
                    for (int i = 0; i < properties.Count; i++)
                    {
                        string val = FormatValue(properties[i].GetValue(item));
                        if (val.Length > maxLen[i])
                        {
                            maxLen[i] = val.Length;
                        }
                    }
                }

                double[] colWidths = new double[properties.Count];
                double totalAvailableWidth = 25.0; // Použitelná šířka A4 na šířku
                double sum = 0;

                for (int i = 0; i < properties.Count; i++)
                {
                    double w = maxLen[i] * 0.22;
                    w = Math.Max(2.0, w);
                    w = Math.Min(8.0, w);
                    colWidths[i] = w;
                    sum += w;
                }

                double scale = totalAvailableWidth / sum;
                for (int i = 0; i < colWidths.Length; i++)
                {
                    colWidths[i] *= scale;
                }

                // Vytvoření tabulky
                MigraDocDoc.Tables.Table table = pdfSection.AddTable();
                table.Borders.Width = 0.5;
                table.Borders.Color = MigraDocDoc.Colors.LightGray;

                for (int i = 0; i < properties.Count; i++)
                {
                    table.AddColumn(MigraDocDoc.Unit.FromCentimeter(colWidths[i]));
                }

                // Záhlaví tabulky (Header)
                MigraDocDoc.Tables.Row headerRow = table.AddRow();
                headerRow.HeadingFormat = true;
                headerRow.Format.Font.Bold = true;
                headerRow.Shading.Color = MigraDocDoc.Colors.LightSteelBlue;

                for (int i = 0; i < properties.Count; i++)
                {
                    headerRow.Cells[i].AddParagraph(GetHeaderName(properties[i]));
                    headerRow.Cells[i].Format.Font.Size = 9;
                }

                // Data tabulky
                for (int r = 0; r < dataSection.Data.Count; r++)
                {
                    MigraDocDoc.Tables.Row row = table.AddRow();
                    if (r % 2 == 1)
                    {
                        row.Shading.Color = MigraDocDoc.Colors.WhiteSmoke;
                    }

                    T item = dataSection.Data[r];

                    for (int c = 0; c < properties.Count; c++)
                    {
                        string value = FormatValue(properties[c].GetValue(item));
                        var p = row.Cells[c].AddParagraph(value);
                        p.Format.Font.Size = 8;
                    }
                    row.Format.KeepTogether = true;
                }

                // Malá mezera za tabulkou
                MigraDocDoc.Paragraph spacing = pdfSection.AddParagraph();
                spacing.Format.SpaceAfter = MigraDocDoc.Unit.FromCentimeter(0.5);
            }

            // Stránkování v patičce
            pdfSection.Footers.Primary.AddParagraph("Vygenerováno: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Format.Font.Size = 8;

            // Vykreslení PDF dokumentu
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true)
            {
                Document = document
            };
            renderer.RenderDocument();

            if (File.Exists(pdfPath) && IsFileLocked(pdfPath))
            {
                Console.WriteLine($"Soubor PDF je uzamčen - nelze uložit: {pdfPath}");
                return;
            }

            renderer.PdfDocument.Save(pdfPath);
            Console.WriteLine($"Hotovo! Sloučený soubor PDF byl uložen do {Path.GetFileName(pdfPath)}");
        }

        private static bool IsFileLocked(string path)
        {
            try
            {
                using FileStream stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        #endregion

        #region Sloučený export do Wordu (DOCX)

        /// <summary>
        /// Exportuje sloučený seznam do jednoho DOCX souboru (Word) pomocí OpenXML.
        /// </summary>
        public static void SaveDocxSections<T>(string docxPath, List<ExportSection<T>> sections, string? mainTitle = null, string[]? columns = null) where T : new()
        {
            if (sections == null || sections.Count == 0)
            {
                return;
            }

            if (!Soubory.CanSaveFile(docxPath))
            {
                return;
            }

            List<PropertyInfo> properties = GetPropertiesToExport<T>(columns);

            using WordprocessingDocument wordDoc = WordprocessingDocument.Create(docxPath, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Wordprocessing.Document(new Wordprocessing.Body());
            Wordprocessing.Body body = mainPart.Document.Body!;

            // Orientace na šířku (Landscape)
            body.Append(
                new Wordprocessing.SectionProperties(
                    new Wordprocessing.PageSize()
                    {
                        Width = 16838,   // A4 Landscape v DXA
                        Height = 11906,
                        Orient = Wordprocessing.PageOrientationValues.Landscape
                    },
                    new Wordprocessing.PageMargin()
                    {
                        Top = 720,
                        Right = 720,
                        Bottom = 720,
                        Left = 720
                    }
                )
            );

            // Hlavní nadpis
            if (!string.IsNullOrWhiteSpace(mainTitle))
            {
                body.Append(ParagraphOf(mainTitle, bold: true, fontSizeHalfPoints: 32, spacingAfter: 200));
            }

            foreach (ExportSection<T> dataSection in sections)
            {
                // Nadpis sekce
                body.Append(ParagraphOf(dataSection.Title, bold: true, fontSizeHalfPoints: 24, spacingAfter: 100));

                if (dataSection.Data == null || dataSection.Data.Count == 0)
                {
                    body.Append(ParagraphOf("Žádná data v této sekci.", bold: false, fontSizeHalfPoints: 18, spacingAfter: 150));
                    continue;
                }

                // Tabulka
                Wordprocessing.Table table = new Wordprocessing.Table();
                table.AppendChild(new Wordprocessing.TableProperties(
                    new Wordprocessing.TableStyle { Val = "TableGrid" },
                    new Wordprocessing.TableWidth { Type = Wordprocessing.TableWidthUnitValues.Pct, Width = "5000" },
                    new Wordprocessing.TableLook { Val = "04A0" }
                ));

                // Header řádek
                Wordprocessing.TableRow headerRow = new Wordprocessing.TableRow();
                foreach (PropertyInfo prop in properties)
                {
                    headerRow.Append(Tc(GetHeaderName(prop), header: true, altRow: false));
                }
                table.Append(headerRow);

                // Datové řádky
                for (int i = 0; i < dataSection.Data.Count; i++)
                {
                    T item = dataSection.Data[i];
                    bool alt = (i % 2) == 1;
                    Wordprocessing.TableRow row = new Wordprocessing.TableRow();

                    foreach (PropertyInfo prop in properties)
                    {
                        object? val = prop.GetValue(item);
                        row.Append(Tc(FormatValue(val), header: false, altRow: alt));
                    }
                    table.Append(row);
                }

                body.Append(table);

                // Mezera za tabulkou
                body.Append(new Wordprocessing.Paragraph(new Wordprocessing.Run(new Wordprocessing.Break())));
            }

            mainPart.Document.Save();
            Console.WriteLine($"Hotovo! Sloučený soubor DOCX byl uložen do {Path.GetFileName(docxPath)}");
        }

        private static Wordprocessing.TableCell Tc(string text, bool header, bool altRow)
        {
            string? bg = header ? "E6E6E6" : (altRow ? "F7F7F7" : null);

            Wordprocessing.TableCellProperties props = new Wordprocessing.TableCellProperties(
                new Wordprocessing.TableCellWidth { Type = Wordprocessing.TableWidthUnitValues.Auto },
                new Wordprocessing.TableCellVerticalAlignment { Val = Wordprocessing.TableVerticalAlignmentValues.Center },
                new Wordprocessing.TableCellMargin(
                    new Wordprocessing.LeftMargin { Width = "120", Type = Wordprocessing.TableWidthUnitValues.Dxa },
                    new Wordprocessing.RightMargin { Width = "120", Type = Wordprocessing.TableWidthUnitValues.Dxa },
                    new Wordprocessing.TopMargin { Width = "80", Type = Wordprocessing.TableWidthUnitValues.Dxa },
                    new Wordprocessing.BottomMargin { Width = "80", Type = Wordprocessing.TableWidthUnitValues.Dxa }
                )
            );

            if (!string.IsNullOrWhiteSpace(bg))
            {
                props.Append(new Wordprocessing.Shading
                {
                    Val = Wordprocessing.ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = bg
                });
            }

            Wordprocessing.RunProperties runProps = new Wordprocessing.RunProperties();
            if (header)
            {
                runProps.Append(new Wordprocessing.Bold());
            }
            runProps.Append(new Wordprocessing.FontSize { Val = "20" }); // Velikost písma 10pt (20 half-points)

            Wordprocessing.Paragraph paragraph = new Wordprocessing.Paragraph(
                new Wordprocessing.ParagraphProperties(new Wordprocessing.SpacingBetweenLines { Before = "0", After = "0" }),
                new Wordprocessing.Run(runProps, new Wordprocessing.Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve })
            );

            return new Wordprocessing.TableCell(props, paragraph);
        }

        private static Wordprocessing.Paragraph ParagraphOf(string text, bool bold, int fontSizeHalfPoints, int spacingAfter)
        {
            Wordprocessing.RunProperties runProps = new Wordprocessing.RunProperties(new Wordprocessing.FontSize { Val = fontSizeHalfPoints.ToString() });
            if (bold)
            {
                runProps.Append(new Wordprocessing.Bold());
            }

            return new Wordprocessing.Paragraph(
                new Wordprocessing.ParagraphProperties(new Wordprocessing.SpacingBetweenLines { After = spacingAfter.ToString() }),
                new Wordprocessing.Run(runProps, new Wordprocessing.Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve })
            );
        }

        #endregion
    }
}
