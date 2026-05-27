using Aplikace.Tridy;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;

namespace Aplikace.Export
{
    public static class DocxGenerator
    {
        public static void SaveDocxGen<T>(this List<T> data,string docxPath,string? title = null)
        {
            if (data == null || data.Count == 0)
                throw new InvalidOperationException("Seznam je prázdný.");

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

            var outDir = Path.GetDirectoryName(Path.GetFullPath(docxPath));

            if (!string.IsNullOrWhiteSpace(outDir))
                Directory.CreateDirectory(outDir);

            using var wordDoc = WordprocessingDocument.Create(docxPath,WordprocessingDocumentType.Document);
            var mainPart = wordDoc.AddMainDocumentPart();

            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            var heading = string.IsNullOrWhiteSpace(title) ? $"Přehled {typeof(T).Name}" : title;

            // Nadpis
            //body.Append(ParagraphOf(heading,bold: true,fontSizeHalfPoints: 32,spacingAfter: 120),ParagraphOf(
            //        DateTime.Now.ToString("dd.MM.yyyy HH:mm"),
            //        bold: false,
            //        fontSizeHalfPoints: 20,
            //        spacingAfter: 200)
            //);
            body.Append(
                new SectionProperties(new PageSize()
                    {
                        Width = 16838,   // A4 landscape
                        Height = 11906,
                        Orient = PageOrientationValues.Landscape
                    },
                    new PageMargin()
                    {
                        Top = 720,
                        Right = 720,
                        Bottom = 720,
                        Left = 720
                    }
                )
            );

            // Tabulka
            var table = new Table();

            table.AppendChild(new TableProperties(new TableStyle { Val = "TableGrid" },new TableWidth
                {
                    Type = TableWidthUnitValues.Pct,
                    Width = "5000"
                },
                new TableLook { Val = "04A0" }
            ));

            // Header
            var headerRow = new TableRow();

            foreach (var prop in properties)
            {
                headerRow.Append(
                    Tc(prop.Name, header: true, altRow: false)
                );
            }
            table.Append(headerRow);

            // Data
            for (int i = 0; i < data.Count; i++)
            {
                var item = data[i];

                bool alt = (i % 2) == 1;

                var row = new TableRow();

                foreach (var prop in properties)
                {
                    var value = prop.GetValue(item);

                    row.Append(Tc(
                            FormatValue(value),
                            header: false,
                            altRow: alt)
                    );
                }
                table.Append(row);
            }

            body.Append(table);
            mainPart.Document.Save();
            Console.WriteLine($"Hotovo! Soubor DOCX byl uložen do {Path.GetFileName(Informace.Create.SouborElektroJson)}");
        }

        private static TableCell Tc(string text,bool header,bool altRow)
        {
            var bg = header
                ? "E6E6E6"
                : (altRow ? "F7F7F7" : null);

            var props = new TableCellProperties(
                new TableCellWidth
                {
                    Type = TableWidthUnitValues.Auto
                },

                new TableCellVerticalAlignment
                {
                    Val = TableVerticalAlignmentValues.Center
                },

                new TableCellMargin(
                    new LeftMargin
                    {
                        Width = "120",
                        Type = TableWidthUnitValues.Dxa
                    },

                    new RightMargin
                    {
                        Width = "120",
                        Type = TableWidthUnitValues.Dxa
                    },

                    new TopMargin
                    {
                        Width = "80",
                        Type = TableWidthUnitValues.Dxa
                    },

                    new BottomMargin
                    {
                        Width = "80",
                        Type = TableWidthUnitValues.Dxa
                    }
                )
            );

            if (!string.IsNullOrWhiteSpace(bg))
            {
                props.Append(new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = bg
                });
            }

            var runProps = new RunProperties();

            if (header)
                runProps.Append(new Bold());

            runProps.Append(new FontSize { Val = "22" });

            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0"
                    }),

                new Run(
                    runProps,
                    new Text(text ?? "")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })
            );

            return new TableCell(props, paragraph);
        }

        private static Paragraph ParagraphOf(
            string text,
            bool bold,
            int fontSizeHalfPoints,
            int spacingAfter)
        {
            var runProps = new RunProperties(
                new FontSize
                {
                    Val = fontSizeHalfPoints.ToString()
                });

            if (bold)
                runProps.Append(new Bold());

            return new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines
                    {
                        After = spacingAfter.ToString()
                    }),

                new Run(
                    runProps,
                    new Text(text ?? "")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })
            );
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