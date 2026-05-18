using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.CustomUI;

//using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Sdilene
{
    public class Word
    {
        public void SaveDocx<T>(List<T> list, string cesta)
        {
            if (!SouborDelete(cesta)) return;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(cesta, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new Body());

                // Nastavení jazyka na češtinu
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
                settingsPart.Settings.AppendChild(new Languages() { Val = "cs-CZ" });

                var table = new Table();

                // Styl tabulky
                TableProperties tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                        new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 }
                    )
                );
                table.AppendChild(tblProps);

                var props = typeof(T).GetProperties();

                // Hlavička
                var headerRow = new TableRow();
                foreach (var prop in props)
                {
                    var cell = new TableCell(new Paragraph(new Run(new Text(prop.Name))));
                    headerRow.Append(cell);
                }
                table.Append(headerRow);

                // Data
                foreach (var item in list)
                {
                    var row = new TableRow();
                    foreach (var prop in props)
                    {
                        if (prop.Name == "Item") continue;
                        var value = prop.GetValue(item, null)?.ToString() ?? "";
                        var cell = new TableCell(new Paragraph(new Run(new Text(value))));
                        row.Append(cell);
                    }
                    table.Append(row);
                }

                mainPart.Document.Body.Append(table);

                // 🧭 Otočení stránky na šířku
                var sectionProps = new SectionProperties();

                var pageSize = new PageSize
                {
                    Width = 16840,     // 29.7 cm v twips
                    Height = 11900,    // 21 cm v twips
                    Orient = PageOrientationValues.Landscape
                };

                var pageMargin = new PageMargin
                {
                    Top = 1440,    // 2.54 cm
                    Right = 1440,
                    Bottom = 1440,
                    Left = 1440,
                    Header = 720,
                    Footer = 720,
                    Gutter = 0
                };
                sectionProps.Append(pageSize, pageMargin);
                mainPart.Document.Body.Append(sectionProps);
                mainPart.Document.Save();
            }
            Console.WriteLine($"Hotovo! Uloženo do {cesta}");
        }

        private static bool SouborDelete(string cesta)
        {
            if (File.Exists(cesta))
            {
                Console.WriteLine($"Soubor {cesta} již existuje.");
                try
                {
                    File.Delete(cesta);
                }
                catch (Exception)
                {
                    Console.WriteLine($"Soubor {cesta} Se nepodařilo smazat. Konec");
                    return false;
                }
            }
            return true;
        }

      
        public void SaveDocxList<T>(List<T> list, string cesta)
        {
            if (!SouborDelete(cesta)) return;
            using (var wordDoc = WordprocessingDocument.Create(cesta, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new Body());

                var body = mainPart.Document.Body;

                // Jazyk na češtinu
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
                settingsPart.Settings.AppendChild(new Languages() { Val = "cs-CZ" });
                settingsPart.Settings.Save();

                var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

                int counter = 1;

                foreach (var item in list)
                {
                    // Nadpis
                    body.Append(CreateParagraph("Zařízení č.", counter.ToString(), bold: true, underline: true));

                    // Podtržená čára
                    body.Append(CreateParagraph(new string('-', 30), ""));

                    foreach (var prop in props)
                    {
                        if (prop.Name == "Item") continue;
                        var displayAttr = prop.GetCustomAttribute<DisplayAttribute>();
                        string displayName = displayAttr?.Name ?? prop.Name;

                        var value = prop.GetValue(item, null)?.ToString() ?? "";
                        body.Append(CreateParagraph(displayName, value));
                    }

                    // Prázdný řádek mezi zařízeními
                    body.Append(new Paragraph(new Run(new Text(""))));
                    counter++;
                }
                mainPart.Document.Save();
            }
            Console.WriteLine($"Hotovo! Uloženo do {cesta}");
        }

        static Paragraph CreateParagraph(string text, string value, bool bold = false, bool underline = false)
        {
            var runProps = new RunProperties();

            if (bold)
                runProps.Append(new Bold());
            if (underline)
                runProps.Append(new Underline { Val = UnderlineValues.Single });

            //asi celá stránka 210mm je asi 11900twips
            var paragraphProperties = new ParagraphProperties(
                new DocumentFormat.OpenXml.Wordprocessing.Tabs(
                    //new TabStop() { Val = TabStopValues.Left, Position = 1440 }, // 1. tab – 1 palec = 1440 twips
                    //new TabStop() { Val = TabStopValues.Left, Position = 2880 }, // 2. tab
                    new TabStop() { Val = TabStopValues.Left, Position = 4320 },  // 3. tab
                    new TabStop() { Val = TabStopValues.Left, Position = 5760 }  // 4. tab
                )
            );

            //var run = new Run();
            //run.Append(runProps);
            //run.Append(new Text(text),
            //    new Tab(), new Text(":"),
            //    new Tab(), new Tab(), new Tab(),
            //    new Text(value)
            //);
            //run.Append(new Tab());
            // Každý úsek textu a tabulátor jako samostatný Run
            var runList = new List<OpenXmlElement>
            {
                new Run(runProps.CloneNode(true), new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new TabChar()),
                new Run(new Text(":") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new TabChar()),
                //new Run(new TabChar()),
                new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })
            };

            var paragraph = new Paragraph(paragraphProperties);
            paragraph.Append(runList);
            return paragraph;
            //return new Paragraph(runProps,run);
        }


        

        public void SaveDocxListClass<T>(List<T> list, string cesta)
        {
            if (!SouborDelete(cesta)) return;
            using (var wordDoc = WordprocessingDocument.Create(cesta, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new Body());

                var body = mainPart.Document.Body;

                // Jazyk na češtinu
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
                settingsPart.Settings.AppendChild(new Languages() { Val = "cs-CZ" });
                settingsPart.Settings.Save();

                var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

                int counter = 1;

                foreach (var item in list)
                {
                    body.Append(CreateParagraph($"Zařízení č.", counter.ToString(), bold: true, underline: true));
                    body.Append(CreateParagraph("",new string('-', 30)));

                    AppendPropertiesToBody(item, body);
                    counter++;
                }
                mainPart.Document.Save();
            }
            Console.WriteLine($"Hotovo! Uloženo do {cesta}");
        }

        static void AppendPropertiesToBody(object obj, Body body, int indent = 0)
        {
            if (obj == null) return;

            var props = obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in props)
            {
                if (prop.Name == "Item") continue;
                var displayAttr = prop.GetCustomAttribute<DisplayAttribute>();
                string displayName = displayAttr?.Name ?? prop.Name;

                var value = prop.GetValue(obj);
                //var value = prop.GetValue(obj, null)?.ToString() ?? "";
                if (value == null)
                {
                    //body.Append(CreateParagraph($"{Indent(indent)}{displayName}"," ---"));
                    body.Append(CreateParagraph("displayName", " ---"));
                }
                else if (IsSimpleType(prop.PropertyType))
                {
                    var valueText = prop.GetValue(obj, null)?.ToString() ?? "";
                    //body.Append(CreateParagraph($"{Indent(indent)}{displayName}", valueText));
                    body.Append(CreateParagraph(displayName, valueText));
                }
                else
                {
                    // Rekurzivní výpis vnořené třídy
                    body.Append(CreateParagraph(displayName, ""));
                    AppendPropertiesToBody(value, body, indent + 1);
                }
            }

            // Prázdný řádek po skupině
            body.Append(new Paragraph(new Run(new Text(""))));
        }

        static bool IsSimpleType(Type type)
        {
            return type.IsPrimitive ||
                   type.IsEnum ||
                   type == typeof(string) ||
                   type == typeof(decimal) ||
                   type == typeof(DateTime) ||
                   type == typeof(Guid);
        }

        static Run Indent(int level)
        {
            var run = new Run();
            for (int i = 0; i < level; i++)
                run.Append(new TabChar());

            return new Run(
                new Text("")); // každá úroveň = 4 mezery
        }

    }
}
