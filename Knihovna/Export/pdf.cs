﻿//pouze pro generování PDF, vyžaduje Windows
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace Knihovna.Export
{
    public static class PdfGenerator
    {
        public static void SavePdfGenFlat<T>(this IEnumerable<T> data, string pdfPath, string? title = null)
        {
            if (data == null || !data.Any())
                throw new InvalidOperationException("Seznam je prázdný.");

            if (!Soubory.CanSaveFile(pdfPath)) return;

            var heading = string.IsNullOrWhiteSpace(title)
                ? $"Přehled {typeof(T).Name}"
                : title;

            var document = new Document();

            var section = document.AddSection();

            // ======================
            // A4 LANDSCAPE
            // ======================
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;

            section.PageSetup.TopMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.RightMargin = Unit.FromCentimeter(1.5);

            // ======================
            // STYL
            // ======================
            var normal = document.Styles["Normal"];
            normal.Font.Name = "Calibri";
            normal.Font.Size = 9;

            // ======================
            // NADPIS
            // ======================
            var titleParagraph = section.AddParagraph(heading);
            titleParagraph.Format.Font.Size = 16;
            titleParagraph.Format.Font.Bold = true;
            titleParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.3);

            var dateParagraph = section.AddParagraph(DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
            dateParagraph.Format.Font.Size = 8;
            dateParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.5);

            // ======================
            // VLASTNOSTI
            // ======================
            var properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetIndexParameters().Length == 0)
                .ToArray();

            //string[] start =
            //[
            //    "Radek",
            //    "Tag",
            //    "Pocet",
            //    "Popis",
            //    "Menic",
            //    "Prikon",
            //    "BalenaJednotka",
            //    "Pid",
            //    "Pozice",
            //    "Poznamka"


            // ======================
            // VÝPOČET ŠÍŘEK
            // ======================
            double[] maxLen = new double[properties.Length];

            // header
            for (int i = 0; i < properties.Length; i++)
                maxLen[i] = properties[i].Name.Length;

            // data
            foreach (var item in data)
            {
                for (int i = 0; i < properties.Length; i++)
                {
                    var val = FormatValue(properties[i].GetValue(item));
                    if (val.Length > maxLen[i])
                        maxLen[i] = val.Length;
                }
            }

            // převod na cm
            double[] colWidths = new double[properties.Length];

            double totalAvailableWidth = 25.0; // cca A4 landscape usable area

            double sum = 0;

            for (int i = 0; i < properties.Length; i++)
            {
                double w = maxLen[i] * 0.22; // 0.22 cm / znak

                // clamp
                w = Math.Max(2.5, w);
                w = Math.Min(8.0, w);

                colWidths[i] = w;
                sum += w;
            }

            // škálování na stránku
            double scale = totalAvailableWidth / sum;

            for (int i = 0; i < colWidths.Length; i++)
                colWidths[i] *= scale;

            // ======================
            // TABULKA
            // ======================
            var table = section.AddTable();
            table.Borders.Width = 0.5;

            for (int i = 0; i < properties.Length; i++)
            {
                table.AddColumn(Unit.FromCentimeter(colWidths[i]));
            }

            // ======================
            // HEADER (opakování)
            // ======================
            var header = table.AddRow();
            header.HeadingFormat = true;
            header.Format.Font.Bold = true;
            header.Shading.Color = Colors.LightGray;

            for (int i = 0; i < properties.Length; i++)
            {
                header.Cells[i].AddParagraph(properties[i].Name);
                header.Cells[i].Format.Font.Size = 9;
            }

            // ======================
            // DATA
            // ======================
            var dataList = data.ToList();
            for (int r = 0; r < dataList.Count; r++)
            {
                var row = table.AddRow();

                if (r % 2 == 1)
                    row.Shading.Color = Colors.WhiteSmoke;

                var item = dataList[r];

                for (int c = 0; c < properties.Length; c++)
                {
                    var value = FormatValue(properties[c].GetValue(item));
                    var p = row.Cells[c].AddParagraph(value);
                    p.Format.Font.Size = 8;
                }

                row.Format.KeepTogether = true;
            }

            // ======================
            // FOOTER
            // ======================
            section.Footers.Primary.AddParagraph(
                "Vygenerováno: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm")
            ).Format.Font.Size = 8;

            // ======================
            // RENDER
            // ======================
            var renderer = new PdfDocumentRenderer(true)
            {
                Document = document
            };

            renderer.RenderDocument();

            var dir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!string.IsNullOrWhiteSpace(dir))
                Directory.CreateDirectory(dir);

            if (File.Exists(pdfPath) && IsFileLocked(pdfPath))
            {
                Console.WriteLine($"Soubor je otevřen - NEJDE ULOŽIT: {Path.GetFileName(pdfPath)}");
                return;
            }

            renderer.PdfDocument.Save(pdfPath);
            Console.WriteLine($"Soubor : {Path.GetFileName(pdfPath)} Uložen.");
        }


        public static void SavePdfGen<T>(this IEnumerable<T> data, string pdfPath, string? title = null, string[] columns = null)
        {
            if (data == null || !data.Any())
                throw new InvalidOperationException("Seznam je prázdný.");

            var heading = string.IsNullOrWhiteSpace(title)
                ? $"Přehled {typeof(T).Name}"
                : title;

            var document = new Document();

            var section = document.AddSection();

            // ======================
            // A4 LANDSCAPE
            // ======================
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;

            section.PageSetup.TopMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.RightMargin = Unit.FromCentimeter(1.5);

            // ======================
            // STYL
            // ======================
            var normal = document.Styles["Normal"];
            normal.Font.Name = "Calibri";
            normal.Font.Size = 9;

            // ======================
            // NADPIS
            // ======================
            var titleParagraph = section.AddParagraph(heading);
            titleParagraph.Format.Font.Size = 16;
            titleParagraph.Format.Font.Bold = true;
            titleParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.3);

            var dateParagraph = section.AddParagraph(DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
            dateParagraph.Format.Font.Size = 8;
            dateParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.5);

            // ======================
            // VLASTNOSTI
            // ======================
            var allProperties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetIndexParameters().Length == 0)
                .ToArray();

            PropertyInfo[] properties;
            if (columns != null && columns.Length > 0)
            {
                properties = columns
                    .Select(colName => allProperties.FirstOrDefault(p => string.Equals(p.Name, colName, StringComparison.OrdinalIgnoreCase)))
                    .Where(p => p != null)
                    .ToArray();
            }
            else
            {
                properties = allProperties;
            }

            // ======================
            // VÝPOČET ŠÍŘEK
            // ======================
            double[] maxLen = new double[properties.Length];

            // header
            for (int i = 0; i < properties.Length; i++)
                maxLen[i] = properties[i].Name.Length;

            // data
            foreach (var item in data)
            {
                for (int i = 0; i < properties.Length; i++)
                {
                    var val = FormatValue(properties[i].GetValue(item));
                    if (val.Length > maxLen[i])
                        maxLen[i] = val.Length;
                }
            }

            // převod na cm
            double[] colWidths = new double[properties.Length];

            double totalAvailableWidth = 25.0; // cca A4 landscape usable area

            double sum = 0;

            for (int i = 0; i < properties.Length; i++)
            {
                double w = maxLen[i] * 0.22; // 0.22 cm / znak

                // clamp
                w = Math.Max(2.5, w);
                w = Math.Min(8.0, w);

                colWidths[i] = w;
                sum += w;
            }

            // škálování na stránku
            double scale = totalAvailableWidth / sum;

            for (int i = 0; i < colWidths.Length; i++)
                colWidths[i] *= scale;

            // ======================
            // TABULKA
            // ======================
            var table = section.AddTable();
            table.Borders.Width = 0.5;

            for (int i = 0; i < properties.Length; i++)
            {
                table.AddColumn(Unit.FromCentimeter(colWidths[i]));
            }

            // ======================
            // HEADER (opakování)
            // ======================
            var header = table.AddRow();
            header.HeadingFormat = true;
            header.Format.Font.Bold = true;
            header.Shading.Color = Colors.LightGray;

            for (int i = 0; i < properties.Length; i++)
            {
                var displayAttr = properties[i].GetCustomAttribute<DisplayAttribute>();
                var name = displayAttr?.Name ?? properties[i].Name;
                var unit = displayAttr?.Prompt; // použijeme jako jednotku
                var nadpis = string.IsNullOrEmpty(unit) ? name
                    : $"{name} [{unit}]";
                header.Cells[i].AddParagraph(nadpis);
                header.Cells[i].Format.Font.Size = 9;
            }

            // ======================
            // DATA
            // ======================
            var dataList = data.ToList();
            for (int r = 0; r < dataList.Count; r++)
            {
                var row = table.AddRow();

                if (r % 2 == 1)
                    row.Shading.Color = Colors.WhiteSmoke;

                var item = dataList[r];

                for (int c = 0; c < properties.Length; c++)
                {
                    var value = FormatValue(properties[c].GetValue(item));
                    var p = row.Cells[c].AddParagraph(value);
                    p.Format.Font.Size = 8;
                }

                row.Format.KeepTogether = true;
            }

            // ======================
            // FOOTER
            // ======================
            section.Footers.Primary.AddParagraph(
                "Vygenerováno: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm")
            ).Format.Font.Size = 8;

            // ======================
            // RENDER
            // ======================
            var renderer = new PdfDocumentRenderer(true)
            {
                Document = document
            };

            renderer.RenderDocument();

            var dir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!string.IsNullOrWhiteSpace(dir))
                Directory.CreateDirectory(dir);

            if (File.Exists(pdfPath) && IsFileLocked(pdfPath))
            {
                Console.WriteLine($"Soubor je otevřen - NEJDE ULOŽIT: {Path.GetFileName(pdfPath)}");
                return;
            }

            renderer.PdfDocument.Save(pdfPath);
            Console.WriteLine($"Hotovo: {Path.GetFileName(pdfPath)}");
        }

        private static string FormatValue(object? value)
        {
            if (value == null)
                return "";

            return value switch
            {
                DateTime dt => dt.ToString("dd.MM.yyyy HH:mm:ss"),
                bool b => b ? "Ano" : "Ne",
                _ => value.ToString() ?? ""
            };
        }

        private static bool IsFileLocked(string path)
        {
            try
            {
                using var stream = new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.ReadWrite,
                    FileShare.None);

                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }
    }
}
