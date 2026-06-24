using System.ComponentModel.DataAnnotations;
using System.Reflection;
using Knihovna.Tridy;

namespace WinForms
{
    public sealed class FormCesty : Form
    {
        private readonly Informace _info;

        // Ovládací prvky pro cesty
        private readonly TextBox _tbBasePath = new() { Dock = DockStyle.Fill };
        private readonly TextBox _tbSouborStrojeXls = new() { Dock = DockStyle.Fill };
        private readonly TextBox _tbSouborStrojeJson = new() { Dock = DockStyle.Fill };
        private readonly TextBox _tbSouborElektroJson = new() { Dock = DockStyle.Fill };
        private readonly TextBox _tbAdresarZdrojDat = new() { Dock = DockStyle.Fill };

        // Popisky stavu pro živou kontrolu
        private readonly Label _lblStatusBasePath = new() { AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Padding = new Padding(6, 6, 0, 0) };
        private readonly Label _lblStatusSouborStrojeXls = new() { AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Padding = new Padding(6, 6, 0, 0) };
        private readonly Label _lblStatusSouborStrojeJson = new() { AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Padding = new Padding(6, 6, 0, 0) };
        private readonly Label _lblStatusSouborElektroJson = new() { AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Padding = new Padding(6, 6, 0, 0) };
        private readonly Label _lblStatusAdresarZdrojDat = new() { AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Padding = new Padding(6, 6, 0, 0) };

        public FormCesty()
        {
            Text = "Kontrola a nastavení cest";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.Sizable;
            AutoScaleMode = AutoScaleMode.Font;

            _info = Informace.Instance;

            BuildUi();
            LoadFromInfo();
            SetupLiveValidation();
        }

        private void BuildUi()
        {
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 7,
                Padding = new Padding(12),
                AutoScroll = true
            };

            // Definice sloupců:
            // 1. Sloupec: Popisek cesty (AutoSize)
            // 2. Sloupec: TextBox s cestou (Percent, 100%)
            // 3. Sloupec: Tlačítko pro procházení "..." (AutoSize)
            // 4. Sloupec: Indikátor stavu (AutoSize)
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180F)); // Pevná šířka pro indikátor stavu, aby neskákala šířka formuláře

            int row = 0;
            AddRow(table, ref row, nameof(Informace.BasePath), _tbBasePath, CreateBrowseFolderButton(_tbBasePath), _lblStatusBasePath);
            AddRow(table, ref row, nameof(Informace.SouborStrojeXls), _tbSouborStrojeXls, CreateBrowseFileButton(_tbSouborStrojeXls, "Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Všechny soubory (*.*)|*.*"), _lblStatusSouborStrojeXls);
            AddRow(table, ref row, nameof(Informace.SouborStrojeJson), _tbSouborStrojeJson, CreateBrowseFileButton(_tbSouborStrojeJson, "JSON (*.json)|*.json|Všechny soubory (*.*)|*.*"), _lblStatusSouborStrojeJson);
            AddRow(table, ref row, nameof(Informace.SouborElektroJson), _tbSouborElektroJson, CreateBrowseFileButton(_tbSouborElektroJson, "JSON (*.json)|*.json|Všechny soubory (*.*)|*.*"), _lblStatusSouborElektroJson);
            AddRow(table, ref row, nameof(Informace.AdresarZdrojDat), _tbAdresarZdrojDat, CreateBrowseFolderButton(_tbAdresarZdrojDat), _lblStatusAdresarZdrojDat);

            // Oddělovač
            table.RowCount = row + 1;
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 10F));
            row++;

            // Panel s tlačítky dole
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                WrapContents = false,
            };

            var btnZrusit = new Button { Text = "Storno", AutoSize = true, DialogResult = DialogResult.Cancel };
            btnZrusit.Click += (_, _) => Close();

            var btnUlozit = new Button { Text = "Uložit nastavení", AutoSize = true, Font = new Font(Font, FontStyle.Bold) };
            btnUlozit.Click += (_, _) => SaveAndClose();

            var btnZkontrolovat = new Button { Text = "Zkontrolovat cesty znovu", AutoSize = true };
            btnZkontrolovat.Click += (_, _) => CheckAllPathsWithReport();

            buttons.Controls.Add(btnZrusit);
            buttons.Controls.Add(btnUlozit);
            buttons.Controls.Add(btnZkontrolovat);

            table.RowCount = row + 1;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(buttons, 0, row);
            table.SetColumnSpan(buttons, 4);

            AcceptButton = btnUlozit;
            CancelButton = btnZrusit;

            Controls.Add(table);
            MinimumSize = new System.Drawing.Size(850, 320);
            Size = new System.Drawing.Size(900, 320);
        }

        private static void AddRow(TableLayoutPanel table, ref int row, string propertyName, TextBox editor, Control browseButton, Label statusLabel)
        {
            table.RowCount = row + 1;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var label = new Label
            {
                Text = GetDisplayName(typeof(Informace), propertyName) ?? propertyName,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Padding = new Padding(0, 6, 12, 0),
            };

            editor.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            browseButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            table.Controls.Add(label, 0, row);
            table.Controls.Add(editor, 1, row);
            table.Controls.Add(browseButton, 2, row);
            table.Controls.Add(statusLabel, 3, row);

            row++;
        }

        private static string? GetDisplayName(Type type, string propertyName)
        {
            var prop = type.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            if (prop is null) return null;

            var display = prop.GetCustomAttributes(typeof(DisplayAttribute), inherit: true).OfType<DisplayAttribute>().FirstOrDefault();
            return display?.Name;
        }

        private static Button CreateBrowseFolderButton(TextBox target)
        {
            var btn = new Button { Text = "...", Width = 36, Height = 24 };
            btn.Click += (_, _) =>
            {
                using var dlg = new FolderBrowserDialog
                {
                    Description = "Vyberte složku",
                    UseDescriptionForTitle = true,
                };

                if (Directory.Exists(target.Text))
                    dlg.SelectedPath = target.Text;

                if (dlg.ShowDialog() == DialogResult.OK)
                    target.Text = dlg.SelectedPath;
            };
            return btn;
        }

        private static Button CreateBrowseFileButton(TextBox target, string filter)
        {
            var btn = new Button { Text = "...", Width = 36, Height = 24 };
            btn.Click += (_, _) =>
            {
                using var dlg = new OpenFileDialog
                {
                    Filter = filter,
                    CheckFileExists = false,
                    RestoreDirectory = true,
                };

                try
                {
                    var dir = Path.GetDirectoryName(target.Text);
                    if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir))
                        dlg.InitialDirectory = dir;
                }
                catch
                {
                    // Ignorovat neplatné cesty
                }

                if (dlg.ShowDialog() == DialogResult.OK)
                    target.Text = dlg.FileName;
            };
            return btn;
        }

        private void LoadFromInfo()
        {
            _tbBasePath.Text = _info.BasePath ?? string.Empty;
            _tbSouborStrojeXls.Text = _info.SouborStrojeXls ?? string.Empty;
            _tbSouborStrojeJson.Text = _info.SouborStrojeJson ?? string.Empty;
            _tbSouborElektroJson.Text = _info.SouborElektroJson ?? string.Empty;
            _tbAdresarZdrojDat.Text = _info.AdresarZdrojDat ?? string.Empty;
        }

        private void SetupLiveValidation()
        {
            // Připojení událostí pro živou validaci
            _tbBasePath.TextChanged += (_, _) => UpdatePathStatus(_tbBasePath, _lblStatusBasePath, isFolder: true);
            _tbSouborStrojeXls.TextChanged += (_, _) => UpdatePathStatus(_tbSouborStrojeXls, _lblStatusSouborStrojeXls, isFolder: false);
            _tbSouborStrojeJson.TextChanged += (_, _) => UpdatePathStatus(_tbSouborStrojeJson, _lblStatusSouborStrojeJson, isFolder: false);
            _tbSouborElektroJson.TextChanged += (_, _) => UpdatePathStatus(_tbSouborElektroJson, _lblStatusSouborElektroJson, isFolder: false);
            _tbAdresarZdrojDat.TextChanged += (_, _) => UpdatePathStatus(_tbAdresarZdrojDat, _lblStatusAdresarZdrojDat, isFolder: true);

            // Prvotní kontrola při načtení
            UpdatePathStatus(_tbBasePath, _lblStatusBasePath, isFolder: true);
            UpdatePathStatus(_tbSouborStrojeXls, _lblStatusSouborStrojeXls, isFolder: false);
            UpdatePathStatus(_tbSouborStrojeJson, _lblStatusSouborStrojeJson, isFolder: false);
            UpdatePathStatus(_tbSouborElektroJson, _lblStatusSouborElektroJson, isFolder: false);
            UpdatePathStatus(_tbAdresarZdrojDat, _lblStatusAdresarZdrojDat, isFolder: true);
        }

        private static void UpdatePathStatus(TextBox tb, Label lbl, bool isFolder)
        {
            string path = tb.Text.Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                lbl.Text = "⚠ Cesta není zadána";
                lbl.ForeColor = Color.DarkGoldenrod;
                return;
            }

            try
            {
                bool exists = isFolder ? Directory.Exists(path) : File.Exists(path);
                if (exists)
                {
                    lbl.Text = "✔ Existuje";
                    lbl.ForeColor = Color.Green;
                }
                else
                {
                    lbl.Text = "✖ Neexistuje";
                    lbl.ForeColor = Color.Red;
                }
            }
            catch
            {
                lbl.Text = "✖ Chyba cesty";
                lbl.ForeColor = Color.Red;
            }
        }

        private void CheckAllPathsWithReport()
        {
            // Provedeme okamžitou aktualizaci
            UpdatePathStatus(_tbBasePath, _lblStatusBasePath, isFolder: true);
            UpdatePathStatus(_tbSouborStrojeXls, _lblStatusSouborStrojeXls, isFolder: false);
            UpdatePathStatus(_tbSouborStrojeJson, _lblStatusSouborStrojeJson, isFolder: false);
            UpdatePathStatus(_tbSouborElektroJson, _lblStatusSouborElektroJson, isFolder: false);
            UpdatePathStatus(_tbAdresarZdrojDat, _lblStatusAdresarZdrojDat, isFolder: true);

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Výsledek kontroly nastavených cest:");
            sb.AppendLine();

            AppendReportLine(sb, "Základní složka projektu", _tbBasePath.Text, isFolder: true);
            AppendReportLine(sb, "Základní soubor strojů", _tbSouborStrojeXls.Text, isFolder: false);
            AppendReportLine(sb, "Stroje JSON", _tbSouborStrojeJson.Text, isFolder: false);
            AppendReportLine(sb, "Elektro JSON", _tbSouborElektroJson.Text, isFolder: false);
            AppendReportLine(sb, "Zdroj dat", _tbAdresarZdrojDat.Text, isFolder: true);

            MessageBox.Show(sb.ToString(), "Kontrola cest", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void AppendReportLine(System.Text.StringBuilder sb, string name, string path, bool isFolder)
        {
            sb.Append(name).Append(": ");
            if (string.IsNullOrWhiteSpace(path))
            {
                sb.AppendLine("[NENÍ ZADÁNO]");
            }
            else
            {
                try
                {
                    bool exists = isFolder ? Directory.Exists(path) : File.Exists(path);
                    sb.AppendLine(exists ? "OK" : "NEEXISTUJE");
                }
                catch
                {
                    sb.AppendLine("CHYBNÁ CESTA");
                }
            }
        }

        private void ApplyToInfo()
        {
            _info.BasePath = _tbBasePath.Text.Trim();
            _info.SouborStrojeXls = _tbSouborStrojeXls.Text.Trim();
            _info.SouborStrojeJson = _tbSouborStrojeJson.Text.Trim();
            _info.SouborElektroJson = _tbSouborElektroJson.Text.Trim();
            _info.AdresarZdrojDat = _tbAdresarZdrojDat.Text.Trim();
        }

        private void SaveAndClose()
        {
            ApplyToInfo();
            _info.Ulozit();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
