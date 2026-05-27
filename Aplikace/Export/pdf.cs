using Aplikace.Sdilene;
using Aplikace.Tridy;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using System.Reflection;

namespace Aplikace.Export
{
    public static class PdfGenerator
    {
        public static void SavePdfGen<T>(this List<T> data, string pdfPath, string? title = null)
        {
            if (data == null || data.Count == 0)
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
            var properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetIndexParameters().Length == 0)
                .ToArray();

            string[] start =
            [
                "Radek",
                "Tag",
                "Pocet",
                "Popis",
                "Menic",
                "Prikon",
                "BalenaJednotka",
                "Pid",
                "Pozice",
                "Poznamka"
            ];

            properties = properties.Where(p => start.Contains(p.Name)).ToArray();

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
            for (int r = 0; r < data.Count; r++)
            {
                var row = table.AddRow();

                if (r % 2 == 1)
                    row.Shading.Color = Colors.WhiteSmoke;

                var item = data[r];

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
                //throw new InvalidOperationException("PDF je otevřený v jiném programu. Zavři ho a zkus znovu.");
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

//using Aplikace.Sdilene;
//using Aplikace.Tridy;
//using MigraDoc.DocumentObjectModel;
//using MigraDoc.DocumentObjectModel.Tables;
//using MigraDoc.Rendering;
//using System.Reflection;

//namespace Aplikace.Export
//{
//    public static class PdfGenerator
//    {
//        //
//        private static double MeasureTextWidth(string text, string fontName = "Calibri", double fontSize = 9)
//        {
//            var font = new System.Drawing.Font(fontName, (float)fontSize);

//            using var bmp = new Bitmap(1, 1);
//            using var g = Graphics.FromImage(bmp);

//            var size = g.MeasureString(text, font);

//            return size.Width; // v pixelech
//        }

//        public static void SavePdfGen<T>(this List<T> data, string pdfPath, string? title = null)
//        {
//            if (data == null || data.Count == 0)
//                throw new InvalidOperationException("Seznam je prázdný.");

//            var heading = string.IsNullOrWhiteSpace(title)
//                ? $"Přehled {typeof(T).Name}"
//                : title;

//            var document = new Document();

//            // ======================
//            // STYLY
//            // ======================
//            var normal = document.Styles["Normal"];
//            normal.Font.Name = "Calibri";
//            normal.Font.Size = 9;

//            var headerStyle = document.Styles.AddStyle("HeaderStyle", "Normal");
//            headerStyle.Font.Bold = true;
//            headerStyle.Font.Size = 9;

//            // ======================
//            // SEKCE
//            // ======================
//            var section = document.AddSection();

//            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;

//            section.PageSetup.PageFormat = PageFormat.A4;

//            section.PageSetup.TopMargin = Unit.FromCentimeter(1.5);
//            section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
//            section.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
//            section.PageSetup.RightMargin = Unit.FromCentimeter(1.5);

//            // ======================
//            // NADPIS
//            // ======================
//            var titleParagraph = section.AddParagraph(heading);
//            titleParagraph.Format.Font.Size = 16;
//            titleParagraph.Format.Font.Bold = true;
//            titleParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.3);

//            var dateParagraph = section.AddParagraph(DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
//            dateParagraph.Format.Font.Size = 8;
//            dateParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.5);

//            // ======================
//            // PROPERTIES
//            // ======================
//            var properties = typeof(T)
//                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
//                .Where(p => p.GetIndexParameters().Length == 0)
//                .ToArray();

//            string[] start =
//            [
//                "Radek",
//                "Tag",
//                "Pocet",
//                "Popis",
//                "Menic",
//                "Prikon",
//                "BalenaJednotka",
//                "Pid",
//                "Pozice",
//                "Poznamka"
//            ];

//            properties = properties.Where(p => start.Contains(p.Name)).ToArray();

//            // ======================
//            // TABULKA
//            // ======================
//            var table = section.AddTable();
//            table.Borders.Width = 0.5;

//            table.Rows.LeftIndent = 0;

//            // 🔥 DŮLEŽITÉ: tabulka přes celou šířku stránky
//            double pageWidth = section.PageSetup.PageWidth
//                               - section.PageSetup.LeftMargin
//                               - section.PageSetup.RightMargin;

//            // ======================
//            // SLUPCE (dynamické šířky)
//            // ======================
//            foreach (var prop in properties)
//            {
//                double widthCm = prop.Name switch
//                {
//                    "Popis" => 5,
//                    "Poznamka" => 6,
//                    "Tag" => 3,
//                    "Prikon" => 2.5,
//                    "Pocet" => 2,
//                    _ => 3
//                };

//                table.AddColumn(Unit.FromCentimeter(widthCm));
//            }

//            //table.SetEdge(0, 0, properties.Length, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.75, Colors.Black);
//            //table.SetEdge(0,0,table.Rows.Count,table.Columns.Count,Edge.Box,MigraDoc.DocumentObjectModel.BorderStyle.Single,0.75,Colors.Black);
//            table.SetEdge(0, 0, table.Columns.Count, table.Rows.Count, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 0.75, Colors.Black);
//            // ======================
//            // HEADER (opakování na stránkách)
//            // ======================
//            var header = table.AddRow();
//            header.HeadingFormat = true;   // 🔥 KLÍČOVÉ
//            header.Format.Font.Bold = true;
//            header.Shading.Color = Colors.LightGray;

//            for (int i = 0; i < properties.Length; i++)
//            {
//                header.Cells[i].AddParagraph(properties[i].Name);
//                header.Cells[i].Format.Font.Size = 9;
//            }

//            // ======================
//            // DATA
//            // ======================
//            for (int r = 0; r < data.Count; r++)
//            {
//                var row = table.AddRow();

//                if (r % 2 == 1)
//                    row.Shading.Color = Colors.WhiteSmoke;

//                var item = data[r];

//                for (int c = 0; c < properties.Length; c++)
//                {
//                    var value = properties[c].GetValue(item);
//                    var cell = row.Cells[c];

//                    var p = cell.AddParagraph(FormatValue(value));
//                    p.Format.Font.Size = 8;   // 🔥 menší font pro data
//                }

//                // 🔥 zabrání rozdělení řádku mezi stránky (volitelné)
//                row.Format.KeepTogether = true;
//            }

//            // ======================
//            // FOOTER
//            // ======================
//            section.Footers.Primary.AddParagraph(
//                "Vygenerováno: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm")
//            ).Format.Font.Size = 8;

//            // ======================
//            // RENDER
//            // ======================
//            var renderer = new PdfDocumentRenderer(true)
//            {
//                Document = document
//            };

//            renderer.RenderDocument();

//            var pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
//            if (!string.IsNullOrWhiteSpace(pdfDir))
//                Directory.CreateDirectory(pdfDir);

//            renderer.PdfDocument.Save(pdfPath);

//            Console.WriteLine($"Hotovo! PDF uloženo: {pdfPath}");
//        }

//        private static string FormatValue(object? value)
//        {
//            if (value == null)
//                return "";

//            return value switch
//            {
//                DateTime dt => dt.ToString("dd.MM.yyyy HH:mm:ss"),
//                bool b => b ? "Ano" : "Ne",
//                _ => value.ToString() ?? ""
//            };
//        }
//    }
//}


//using Aplikace.Sdilene;
//using Aplikace.Tridy;
//using MigraDoc.DocumentObjectModel;
//using MigraDoc.Rendering;
//using System.Reflection;

//namespace Aplikace.Export
//{
//    public static class PdfGenerator
//    {
//        public static void SavePdfGen<T>(this List<T> data, string pdfPath, string? title = null)
//        {
//            if (data == null || data.Count == 0)
//                throw new InvalidOperationException("Seznam je prázdný.");

//            var heading = string.IsNullOrWhiteSpace(title) ? $"Přehled {typeof(T).Name}" : title;

//            var document = new Document();

//            var section = document.AddSection();

//            // Okraje
//            section.PageSetup.TopMargin = Unit.FromCentimeter(2);
//            section.PageSetup.BottomMargin = Unit.FromCentimeter(2);
//            section.PageSetup.LeftMargin = Unit.FromCentimeter(2);
//            section.PageSetup.RightMargin = Unit.FromCentimeter(2);

//            // Styl
//            var style = document.Styles["Normal"];
//            style.Font.Name = "Calibri";
//            style.Font.Size = 10;

//            // Nadpis
//            var titleParagraph = section.AddParagraph();
//            titleParagraph.AddFormattedText(heading, TextFormat.Bold);
//            titleParagraph.Format.Font.Size = 18;
//            titleParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.5);

//            // Datum
//            var dateParagraph = section.AddParagraph(DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
//            dateParagraph.Format.Font.Size = 9;
//            dateParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.5);

//            // Reflection - vlastnosti
//            //var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
//            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(p => p.GetIndexParameters().Length == 0).ToArray();

//            //filtr vlastností - pouze ty, které jsou v seznamu start
//            string[] start = [
//                "Radek",
//                "Tag",
//                "Pocet",
//                "Popis",
//                "Menic",
//                "Prikon",
//                "BalenaJednotka",
//                "Pid",
//                "Pozice",
//                "Poznamka" ];
//            properties = [.. properties.Where(p => start.Contains(p.Name))];

//            // Tabulka
//            var table = section.AddTable();
//            table.Borders.Width = 0.5;

//            // Sloupce
//            foreach (var prop in properties)
//            {
//                table.AddColumn(Unit.FromCentimeter(4));
//            }

//            // Header
//            var header = table.AddRow();
//            header.Shading.Color = Colors.LightGray;
//            header.Format.Font.Bold = true;

//            for (int i = 0; i < properties.Length; i++)
//            {
//                header.Cells[i].AddParagraph(properties[i].Name);
//            }

//            // Data
//            for (int r = 0; r < data.Count; r++)
//            {
//                var row = table.AddRow();

//                if (r % 2 == 1)
//                    row.Shading.Color = Colors.WhiteSmoke;

//                var item = data[r];

//                for (int c = 0; c < properties.Length; c++)
//                {
//                    var value = properties[c].GetValue(item);

//                    row.Cells[c].AddParagraph(FormatValue(value));
//                }
//            }

//            // Footer
//            section.Footers.Primary.AddParagraph().AddText("Vygenerováno: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm"));

//            // Render PDF
//            var renderer = new PdfDocumentRenderer(true)
//            {
//                Document = document
//            };

//            renderer.RenderDocument();

//            var pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));

//            if (!string.IsNullOrWhiteSpace(pdfDir))
//                Directory.CreateDirectory(pdfDir);

//            renderer.PdfDocument.Save(pdfPath);
//            //Console.WriteLine($"Hotovo! Uloženo do {pdfPath}");
//            Console.WriteLine($"Hotovo! Soubor PDF byl uložen do {Path.GetFileName(Informace.Create.SouborElektroJson)}");
//        }

//        private static string FormatValue(object? value)
//        {
//            if (value == null)
//                return "";

//            return value switch
//            {
//                DateTime dt => dt.ToString("dd.MM.yyyy HH:mm:ss"),
//                bool b => b ? "Ano" : "Ne",
//                _ => value.ToString() ?? ""
//            };
//        }
//    }
//}