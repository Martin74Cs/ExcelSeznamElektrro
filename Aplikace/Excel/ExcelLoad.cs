using Aplikace.Sdilene;
using Aplikace.Tridy;

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Excel
{
    public class ExcelLoad
    {

        /// <summary> Načtení dokumentu Ecxel nebo Json do pole List<List<string>> z a vytvořejí JSON</summary>
        public static List<List<string>> LoadDataExcel(string cesta, int[] Sloupce, string Tabulka , int Radek)
        {
            Console.Write("\nProbíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            //var Pole = new List<List<string>>();
            //string Soubor = Path.GetFileName(cesta);
            //string Adresar = Path.GetDirectoryName(cesta);
            //string json = Path.Combine(Adresar, Path.ChangeExtension(Soubor, ".json"));
            string json = Path.ChangeExtension(cesta, ".json");
            if (File.Exists(json))
            {
                return Soubory.LoadJsonList<List<string>>(json);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            }
            else
            {
                var ExcelApp = new ExcelApp();
                var Pole = ExcelApp.ExelLoadTable(cesta, Tabulka, Radek, Sloupce);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                if(Pole.Count>1) Pole.SaveJsonList(json);
                Console.WriteLine($"načeno {Pole.Count} záznamů.");
                return Pole;
            }
        }

        /// <summary> Načtení dokumentu Ecxel nebo Json do pole List List string z a vytvořejí JSON</summary>
        public static List<Zarizeni> DataExcel(string cesta, string Tabulka, int Radek)
        {
            Console.WriteLine("Probíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            //var Pole = new List<List<string>>();
            //var Pole = new List<Zarizeni>();
            //string Soubor = Path.GetFileName(cesta);
            //string Adresar = Path.GetDirectoryName(cesta);

            if (!File.Exists(cesta)) return [];

            var ExcelApp = new ExcelApp(cesta);
            //ExcelApp.DokumetExcel(cesta);
            ExcelApp.GetSheet(Tabulka);
            if (ExcelApp.Xls == null) return [];
            
            if (ExcelApp.Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.WriteLine("Sheet=" + ExcelApp.Xls.Name);

            //Sloupce které se mají načíst z Excelu do názvů tříd. Myslím že třída musí existovat
            var dir = new Dictionary<int, string>() {
                {1, "Radek"     },
                {2, "Tag"       },
                {3, "Pocet"     },
                {4, "Popis"     },
                {11, "Menic"    },
                {10, "Prikon"   },
                {18, "BalenaJednotka"   },
            };

            var Pole = ExcelApp.ExelTable(Radek,Tabulka, dir);

            ExcelApp.ExcelQuit(cesta);
            //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            //if (Pole.Count > 1) Pole.SaveJsonList(Cesty.ElektroRozvaděčJson);
            Console.WriteLine($"načeno {Pole.Count} záznamů.");
            return Pole;
        }

        public class ColumnMappingConfig
        {
            public string FilePath { get; set; } = "";
            public string SheetName { get; set; } = "";
            public int StartRow { get; set; } = 8;
            public Dictionary<string, int> PropertyToColumn { get; set; } = new Dictionary<string, int>();
        }

        private static List<ColumnMappingConfig> LoadAllMappings()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string file = Path.Combine(appData, "Elektro", "column_mappings.json");
                if (File.Exists(file))
                {
                    string json = File.ReadAllText(file);
                    return Newtonsoft.Json.JsonConvert.DeserializeObject<List<ColumnMappingConfig>>(json) ?? new List<ColumnMappingConfig>();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chyba při načítání konfigurace mapování sloupců: {ex.Message}");
            }
            return new List<ColumnMappingConfig>();
        }

        private static void SaveMapping(string filePath, string sheetName, int startRow, Dictionary<int, string> mapping)
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string dir = Path.Combine(appData, "Elektro");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                string file = Path.Combine(dir, "column_mappings.json");
                
                var allConfigs = LoadAllMappings();
                
                allConfigs.RemoveAll(x => x.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
                
                var propToCol = new Dictionary<string, int>();
                foreach (var kvp in mapping)
                {
                    propToCol[kvp.Value] = kvp.Key;
                }
                
                allConfigs.Add(new ColumnMappingConfig
                {
                    FilePath = filePath,
                    SheetName = sheetName,
                    StartRow = startRow,
                    PropertyToColumn = propToCol
                });
                
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(allConfigs, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(file, json);
                Console.WriteLine("Konfigurace přiřazení sloupců byla úspěšně uložena.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chyba při ukládání konfigurace mapování sloupců: {ex.Message}");
            }
        }

        private static void FillPreviewGrid(ClosedXML.Excel.IXLWorksheet ws, System.Windows.Forms.DataGridView dgv)
        {
            dgv.Rows.Clear();
            dgv.Columns.Clear();

            int rowCount = Math.Min(ws.LastRowUsed()?.RowNumber() ?? 0, 30);
            int colCount = Math.Min(ws.LastColumnUsed()?.ColumnNumber() ?? 0, 20);

            if (rowCount == 0 || colCount == 0) return;

            // Přidání sloupců s písmeny (A, B, C, D...)
            for (int col = 1; col <= colCount; col++)
            {
                string colLetter = GetExcelColumnName(col);
                dgv.Columns.Add($"Col{col}", $"{colLetter} ({col})");
                dgv.Columns[col - 1].Width = 100;
            }

            // Přidání řádků
            for (int r = 1; r <= rowCount; r++)
            {
                var rowValues = new string[colCount];
                for (int c = 1; c <= colCount; c++)
                {
                    rowValues[c - 1] = ws.Cell(r, c).GetString();
                }
                dgv.Rows.Add(rowValues);
                dgv.Rows[r - 1].HeaderCell.Value = r.ToString();
            }
            dgv.RowHeadersWidth = 60;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Načtení dokumentu Excel s možností interaktivního přiřazení sloupců pomocí dialogu.
        /// </summary>
        public static List<Zarizeni> DataExcelInteractive(string cesta, string Tabulka, int Radek)
        {
            Console.WriteLine("Probíhá načítání dat... ");
            if (!File.Exists(cesta)) return [];

            var ExcelApp = new ExcelApp(cesta);
            var workbook = new ClosedXML.Excel.XLWorkbook(cesta);
            var sheetNames = workbook.Worksheets.Select(x => x.Name).ToList();

            if (sheetNames.Count == 0)
            {
                Console.WriteLine("Excel neobsahuje žádné listy.");
                workbook.Dispose();
                ExcelApp.ExcelQuit(cesta);
                return [];
            }

            // Slovník pro výsledek
            var dir = new Dictionary<int, string>();
            string finalSheetName = Tabulka;
            int finalStartRow = Radek;
            bool dialogConfirmed = false;

            var allConfigs = LoadAllMappings();
            var existingConfig = allConfigs.FirstOrDefault(x => x.FilePath.Equals(cesta, StringComparison.OrdinalIgnoreCase));

            var thread = new System.Threading.Thread(() =>
            {
                using (var form = new System.Windows.Forms.Form())
                {
                    form.Text = "Přiřazení sloupců Excelu k vlastnostem";
                    form.Width = 1000;
                    form.Height = 700;
                    form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;
                    form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    form.Font = new System.Drawing.Font("Segoe UI", 9F);
                    form.BackColor = System.Drawing.Color.FromArgb(245, 246, 248);

                    // Hlavní panel pro rozdělení (Left = Settings, Right = Preview)
                    var panelLeft = new System.Windows.Forms.Panel()
                    {
                        Location = new System.Drawing.Point(10, 10),
                        Size = new System.Drawing.Size(430, 640),
                        BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                        BackColor = System.Drawing.Color.White
                    };
                    form.Controls.Add(panelLeft);

                    var panelRight = new System.Windows.Forms.Panel()
                    {
                        Location = new System.Drawing.Point(450, 10),
                        Size = new System.Drawing.Size(525, 640),
                        BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                        BackColor = System.Drawing.Color.White
                    };
                    form.Controls.Add(panelRight);

                    // --- LEVÝ PANEL (Settings & Mapping) ---
                    var labelTitle = new System.Windows.Forms.Label()
                    {
                        Text = "Nastavení importu a přiřazení sloupců",
                        Location = new System.Drawing.Point(15, 15),
                        Size = new System.Drawing.Size(400, 20),
                        Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold)
                    };
                    panelLeft.Controls.Add(labelTitle);

                    // Výběr listu (Sheet)
                    var labelSheet = new System.Windows.Forms.Label()
                    {
                        Text = "List (Sheet):",
                        Location = new System.Drawing.Point(15, 55),
                        Size = new System.Drawing.Size(100, 20),
                        Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
                    };
                    panelLeft.Controls.Add(labelSheet);

                    var cbSheet = new System.Windows.Forms.ComboBox()
                    {
                        Location = new System.Drawing.Point(130, 52),
                        Width = 280,
                        DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
                    };
                    foreach (var sName in sheetNames) cbSheet.Items.Add(sName);
                    
                    if (existingConfig != null && sheetNames.Contains(existingConfig.SheetName))
                        cbSheet.SelectedItem = existingConfig.SheetName;
                    else if (sheetNames.Contains(Tabulka))
                        cbSheet.SelectedItem = Tabulka;
                    else
                        cbSheet.SelectedIndex = 0;

                    panelLeft.Controls.Add(cbSheet);

                    // Výběr prvního řádku dat (Radek)
                    var labelRow = new System.Windows.Forms.Label()
                    {
                        Text = "První řádek dat:",
                        Location = new System.Drawing.Point(15, 95),
                        Size = new System.Drawing.Size(100, 20),
                        Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
                    };
                    panelLeft.Controls.Add(labelRow);

                    var nudRow = new System.Windows.Forms.NumericUpDown()
                    {
                        Location = new System.Drawing.Point(130, 93),
                        Width = 80,
                        Minimum = 1,
                        Maximum = 100,
                        Value = existingConfig != null ? existingConfig.StartRow : Radek
                    };
                    panelLeft.Controls.Add(nudRow);

                    var labelMappingTitle = new System.Windows.Forms.Label()
                    {
                        Text = "Přiřazení parametrů:",
                        Location = new System.Drawing.Point(15, 140),
                        Size = new System.Drawing.Size(400, 20),
                        Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold),
                        ForeColor = System.Drawing.Color.Navy
                    };
                    panelLeft.Controls.Add(labelMappingTitle);

                    // Kontroly pro parametry
                    var targetProperties = new string[] {
                        "Radek",
                        "Tag",
                        "Pocet",
                        "Popis",
                        "Menic",
                        "Prikon",
                        "BalenaJednotka"
                    };

                    int startY = 175;
                    int spacingY = 48;
                    var comboBoxes = new Dictionary<string, System.Windows.Forms.ComboBox>();

                    for (int i = 0; i < targetProperties.Length; i++)
                    {
                        string prop = targetProperties[i];

                        var labelProp = new System.Windows.Forms.Label()
                        {
                            Text = prop + ":",
                            Location = new System.Drawing.Point(15, startY + (i * spacingY) + 3),
                            Size = new System.Drawing.Size(110, 20),
                            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
                        };
                        panelLeft.Controls.Add(labelProp);

                        var cb = new System.Windows.Forms.ComboBox()
                        {
                            Location = new System.Drawing.Point(130, startY + (i * spacingY)),
                            Width = 280,
                            DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
                        };
                        panelLeft.Controls.Add(cb);
                        comboBoxes[prop] = cb;
                    }

                    // --- PRAVÝ PANEL (Data Preview) ---
                    var labelPreview = new System.Windows.Forms.Label()
                    {
                        Text = "Náhled dat v listu (prvních 30 řádků):",
                        Location = new System.Drawing.Point(15, 15),
                        Size = new System.Drawing.Size(490, 20),
                        Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold)
                    };
                    panelRight.Controls.Add(labelPreview);

                    var dgvPreview = new System.Windows.Forms.DataGridView()
                    {
                        Location = new System.Drawing.Point(15, 52),
                        Size = new System.Drawing.Size(490, 520),
                        AllowUserToAddRows = false,
                        AllowUserToDeleteRows = false,
                        ReadOnly = true,
                        BackgroundColor = System.Drawing.Color.White,
                        ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
                    };
                    panelRight.Controls.Add(dgvPreview);

                    // Akce při změně listu nebo řádku
                    Action updateSheetAndGrid = () =>
                    {
                        string selectedSheet = cbSheet.SelectedItem?.ToString() ?? "";
                        if (string.IsNullOrEmpty(selectedSheet)) return;

                        var ws = workbook.Worksheet(selectedSheet);
                        FillPreviewGrid(ws, dgvPreview);

                        // Aktualizace ComboBoxů sloupců
                        int dataRow = (int)nudRow.Value;
                        int headerRow = dataRow > 1 ? dataRow - 1 : dataRow;

                        int colCount = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
                        var excelCols = new List<Tuple<int, string>>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            string headerName = ws.Cell(headerRow, col).GetString().Trim();
                            if (string.IsNullOrEmpty(headerName))
                            {
                                headerName = $"Sloupec {col}";
                            }
                            excelCols.Add(Tuple.Create(col, $"{col}: {headerName}"));
                        }

                        // Zjistíme, zda existuje uložení pro tento konkrétní soubor a list
                        var currentMapping = existingConfig?.PropertyToColumn;

                        foreach (var kvp in comboBoxes)
                        {
                            var cb = kvp.Value;
                            cb.Items.Clear();
                            cb.Items.Add(new { Key = 0, Value = "-- Nepřiřazeno --" });

                            int selectIndex = 0;
                            int currentIndex = 1;

                            int savedCol = 0;
                            if (currentMapping != null && currentMapping.TryGetValue(kvp.Key, out int val))
                            {
                                savedCol = val;
                            }

                            foreach (var col in excelCols)
                            {
                                cb.Items.Add(new { Key = col.Item1, Value = col.Item2 });

                                if (savedCol > 0)
                                {
                                    if (col.Item1 == savedCol)
                                    {
                                        selectIndex = currentIndex;
                                    }
                                }
                                else
                                {
                                    // Automatická detekce jako fallback
                                    string colLower = col.Item2.ToLower();
                                    string propLower = kvp.Key.ToLower();
                                    bool matches = false;

                                    if (propLower == "radek" && (colLower.Contains("řádek") || colLower.Contains("radek") || colLower.Contains("row") || colLower.Contains("no."))) matches = true;
                                    else if (propLower == "tag" && (colLower.Contains("tag") || colLower.Contains("označení") || colLower.Contains("oznaceni") || colLower.Contains("stroj") || colLower.Contains("motor"))) matches = true;
                                    else if (propLower == "pocet" && (colLower.Contains("počet") || colLower.Contains("pocet") || colLower.Contains("kus") || colLower.Contains("qty") || colLower.Contains("count"))) matches = true;
                                    else if (propLower == "popis" && (colLower.Contains("popis") || colLower.Contains("description") || colLower.Contains("název") || colLower.Contains("nazev"))) matches = true;
                                    else if (propLower == "menic" && (colLower.Contains("měnič") || colLower.Contains("menic") || colLower.Contains("vsd") || colLower.Contains("frekv"))) matches = true;
                                    else if (propLower == "prikon" && (colLower.Contains("příkon") || colLower.Contains("prikon") || colLower.Contains("výkon") || colLower.Contains("kw") || colLower.Contains("hp"))) matches = true;
                                    else if (propLower == "balenajednotka" && (colLower.Contains("balená") || colLower.Contains("balena") || colLower.Contains("jednotka") || colLower.Contains("package") || colLower.Contains("pack"))) matches = true;

                                    if (matches)
                                    {
                                        selectIndex = currentIndex;
                                    }
                                }
                                currentIndex++;
                            }

                            cb.DisplayMember = "Value";
                            cb.ValueMember = "Key";
                            cb.SelectedIndex = selectIndex;
                        }
                    };

                    cbSheet.SelectedIndexChanged += (s, e) => updateSheetAndGrid();
                    nudRow.ValueChanged += (s, e) => updateSheetAndGrid();

                    // Spustíme první naplnění
                    updateSheetAndGrid();

                    // Tlačítka OK a Storno
                    var buttonOk = new System.Windows.Forms.Button()
                    {
                        Text = "Potvrdit a uložit",
                        Location = new System.Drawing.Point(200, 585),
                        Size = new System.Drawing.Size(140, 35),
                        DialogResult = System.Windows.Forms.DialogResult.OK,
                        BackColor = System.Drawing.Color.FromArgb(40, 167, 69),
                        ForeColor = System.Drawing.Color.White,
                        FlatStyle = System.Windows.Forms.FlatStyle.Flat
                    };
                    buttonOk.FlatAppearance.BorderSize = 0;
                    buttonOk.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
                    panelLeft.Controls.Add(buttonOk);

                    var buttonCancel = new System.Windows.Forms.Button()
                    {
                        Text = "Storno",
                        Location = new System.Drawing.Point(350, 585),
                        Size = new System.Drawing.Size(70, 35),
                        DialogResult = System.Windows.Forms.DialogResult.Cancel,
                        BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                        ForeColor = System.Drawing.Color.White,
                        FlatStyle = System.Windows.Forms.FlatStyle.Flat
                    };
                    buttonCancel.FlatAppearance.BorderSize = 0;
                    panelLeft.Controls.Add(buttonCancel);

                    form.AcceptButton = buttonOk;
                    form.CancelButton = buttonCancel;

                    if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        finalSheetName = cbSheet.SelectedItem?.ToString() ?? Tabulka;
                        finalStartRow = (int)nudRow.Value;

                        foreach (var kvp in comboBoxes)
                        {
                            var selectedItem = kvp.Value.SelectedItem;
                            if (selectedItem != null)
                            {
                                var keyProp = selectedItem.GetType().GetProperty("Key");
                                if (keyProp != null)
                                {
                                    int key = (int)keyProp.GetValue(selectedItem)!;
                                    if (key > 0)
                                    {
                                        dir[key] = kvp.Key;
                                    }
                                }
                            }
                        }
                        dialogConfirmed = true;
                    }
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();

            workbook.Dispose();

            if (!dialogConfirmed || dir.Count == 0)
            {
                Console.WriteLine("Mapování sloupců bylo zrušeno nebo nebyly přiřazeny žádné sloupce.");
                ExcelApp.ExcelQuit(cesta);
                return [];
            }

            // Uložíme konfiguraci do persistenčního úložiště
            SaveMapping(cesta, finalSheetName, finalStartRow, dir);

            Console.WriteLine($"\nZvolený list: {finalSheetName}, První řádek: {finalStartRow}");
            Console.WriteLine("Zvolené mapování sloupců:");
            foreach (var kvp in dir)
            {
                Console.WriteLine($"Sloupec {kvp.Key} -> Parametr {kvp.Value}");
            }

            // Načtení dat pomocí vybraného listu a mapování
            ExcelApp.GetSheet(finalSheetName);
            var Pole = ExcelApp.ExelTable(finalStartRow, finalSheetName, dir);
            ExcelApp.ExcelQuit(cesta);
            Console.WriteLine($"načteno {Pole.Count} záznamů.");
            return Pole;
        }

        /// <summary> Načtení dokumentu Ecxel nebo Json do pole List List string z a vytvořejí JSON</summary>
        public static List<Vykres> DataExcelVykres(string cesta, string Tabulka, int Radek)
        {
            Console.WriteLine("Probíná hačítání dat ... ");
            if (!File.Exists(cesta)) return [];

            var ExcelApp = new ExcelApp(cesta);

            ExcelApp.GetSheet(Tabulka);
            if (ExcelApp.Xls == null) return [];
            
            if (ExcelApp.Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.WriteLine("Sheet=" + ExcelApp.Xls.Name);

            //Sloupce které se mají načíst z Excelu do názvů tříd. Myslím že třída musí existovat
            var dir = new Dictionary<int, string>() {
                {2, "OrientačníčísloB"     },
                {3, "OrientačníčísloC"       },
                {4, "OrientačníčísloD"     },
                {5, "OrientačníčísloE"     },
                {6, "OrientačníčísloF"    },
                {7, "ČísloDokumentu"   },
                {8, "Nazev"   },
                {9, "Revize"   },
                {11, "Popisrevize"   },
                {12, "Cesta"   },
                {13, "ProfesníČislo"   },

            };

            var Pole = ExcelApp.ExelTableVykresy(Radek,Tabulka, dir);

            ExcelApp.ExcelQuit(cesta);
            //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            //if (Pole.Count > 1) Pole.SaveJsonList(Cesty.ElektroRozvaděčJson);
            Console.WriteLine($"načeno {Pole.Count} záznamů.");
            return Pole;
        }

        /// <summary> Načtení dokumentu Ecxel nebo Json do pole List Zarizeni z a vytvořejí JSON</summary>
        public static List<Zarizeni> DwgDataExcel(string cesta, string Tabulka, int Radek)
        {
            Console.WriteLine("Probíná hačítání dat ... ");
            if (!File.Exists(cesta)) return [];

            var ExcelApp = new ExcelApp(cesta);
            ExcelApp.GetSheet(Tabulka);
            if (ExcelApp.Xls == null) return [];
            
            if (ExcelApp.Xls == null) { Console.Write("\nChyba KONEC"); return []; }
            Console.WriteLine("Sheet=" + ExcelApp.Xls.Name);

            //Sloupce které se mají načíst z Excelu do názvů tříd. Myslím že třída musí existovat
            //DOPLNIT SLOUPCE PRO DWG
            var dir = new Dictionary<int, string>() {
                //{1, "Radek"   },
                {6, "Predmet"   },
                {7, "PID"       },
                //{3, "Pocet"   },
                {8, "Popis"     },
                {9, "Druh"      },
                {10, "Typ"      },
                {21, "Tag"      },
                {23, "TagStroj" },
                {24, "Menic"    },
                {26, "Prikon"   },
                {25, "Etapa"    },
                {27, "Patro"    },
                //{18, "BalenaJednotka"   },
            };

            var Pole = ExcelApp.ExelTable(Radek,Tabulka, dir);

            ExcelApp.ExcelQuit(cesta);
            Console.WriteLine($"Načeno {Pole.Count} záznamů.");
            return Pole;
            
        }

        /// <summary> Načtení dokumentu Ecxel do pole Třídy z a vytvořejí JSON</summary>
        public static List<Zarizeni> LoadDataExcelTrida(string cesta, int[] Sloupce, string Tabulka , int Radek, string[] TextPole)
        {
            if (!System.IO.File.Exists(cesta)) return [];
            Console.Write("\nProbíná hačítání dat ... ");
            //začíná sloupcem číslo 1

            //var Pole = new List<Zarizeni>();
            string Soubor = Path.GetFileName(cesta);
            string Adresar = Path.GetDirectoryName(cesta) ?? Environment.SpecialFolder.MyDocuments.ToString();
            string json = Path.Combine(Adresar, Path.ChangeExtension(Soubor, ".json"));
            if (File.Exists(json))
            {
                return Soubory.LoadJsonList<Zarizeni>(json);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
            }
            else
            {
                var ExcelApp = new ExcelApp();
                var Pole = ExcelApp.ExelLoadTableTrida(cesta, Tabulka, Radek, Sloupce, TextPole);
                //Pole = Pole.OrderBy(x => Convert.ToDouble(x[0])).ToList();
                Pole.SaveJsonList(json);
                return Pole;
            }
        }

        public static string Apid(int length = 9)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var random = new Random();
            //return new string(Enumerable.Repeat(chars, length)
            //    .Select(s => s[random.Next(s.Length)]).ToArray());
            return new string([.. Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)])]);
        }
    }
}
