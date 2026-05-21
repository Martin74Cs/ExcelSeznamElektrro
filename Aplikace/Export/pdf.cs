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

            // Okraje
            section.PageSetup.TopMargin = Unit.FromCentimeter(2);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(2);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(2);
            section.PageSetup.RightMargin = Unit.FromCentimeter(2);

            // Styl
            var style = document.Styles["Normal"];
            style.Font.Name = "Calibri";
            style.Font.Size = 10;

            // Nadpis
            var titleParagraph = section.AddParagraph();
            titleParagraph.AddFormattedText(heading, TextFormat.Bold);
            titleParagraph.Format.Font.Size = 18;
            titleParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.5);

            // Datum
            var dateParagraph = section.AddParagraph(DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
            dateParagraph.Format.Font.Size = 9;
            dateParagraph.Format.SpaceAfter = Unit.FromCentimeter(0.5);

            // Reflection - vlastnosti
            //var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(p => p.GetIndexParameters().Length == 0).ToArray();

            //filtr vlastností - pouze ty, které jsou v seznamu start
            string[] start = {
                "Radek",
                "Tag",
                "Pocet",
                "Popis",
                "Menic",
                "Prikon",
                "BalenaJednotka",
                "Pid",
                "Pozice",
                "Poznamka" };
            properties = [.. properties.Where(p => start.Contains(p.Name))];

            // Tabulka
            var table = section.AddTable();
            table.Borders.Width = 0.5;

            // Sloupce
            foreach (var prop in properties)
            {
                table.AddColumn(Unit.FromCentimeter(4));
            }

            // Header
            var header = table.AddRow();
            header.Shading.Color = Colors.LightGray;
            header.Format.Font.Bold = true;

            for (int i = 0; i < properties.Length; i++)
            {
                header.Cells[i].AddParagraph(properties[i].Name);
            }

            // Data
            for (int r = 0; r < data.Count; r++)
            {
                var row = table.AddRow();

                if (r % 2 == 1)
                    row.Shading.Color = Colors.WhiteSmoke;

                var item = data[r];

                for (int c = 0; c < properties.Length; c++)
                {
                    var value = properties[c].GetValue(item);

                    row.Cells[c].AddParagraph(FormatValue(value));
                }
            }

            // Footer
            section.Footers.Primary.AddParagraph()
                .AddText("Vygenerováno: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm"));

            // Render PDF
            var renderer = new PdfDocumentRenderer(true)
            {
                Document = document
            };

            renderer.RenderDocument();

            var pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));

            if (!string.IsNullOrWhiteSpace(pdfDir))
                Directory.CreateDirectory(pdfDir);

            renderer.PdfDocument.Save(pdfPath);
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
    }
}