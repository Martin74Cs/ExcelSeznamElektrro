using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Aplikace.Tridy;

namespace Aplikace.Forms;

// Jednoduchý WinForms formulář pro editaci třídy Informace (singleton).
public sealed class Nastaveni : Form
{
    private readonly Informace _info;

    private readonly TextBox _tbBasePath = new() { Dock = DockStyle.Fill };
    private readonly TextBox _tbSouborStrojeXls = new() { Dock = DockStyle.Fill };
    private readonly TextBox _tbSouborStrojeJson = new() { Dock = DockStyle.Fill };
    private readonly TextBox _tbSouborElektroJson = new() { Dock = DockStyle.Fill };
    //private readonly TextBox _tbSouborData = new() { Dock = DockStyle.Fill };

    private readonly TextBox _tbMistnost = new() { Dock = DockStyle.Fill };
    private readonly TextBox _tbProjekt = new() { Dock = DockStyle.Fill };
    private readonly TextBox _tbNazev = new() { Dock = DockStyle.Fill };
    private readonly TextBox _tbPoznamka = new() { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, Height = 80 };
    private readonly TextBox _tbData = new() { Dock = DockStyle.Fill };

    private readonly DateTimePicker _dpDatum = new() { Dock = DockStyle.Left, Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "dd.MM.yyyy HH:mm" };

    public Nastaveni()
    {
        Text = "Nastavení";
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.Sizable;
        AutoScaleMode = AutoScaleMode.Font;

        using var info = Informace.Create;
        _info = info;

        BuildUi();
        LoadFromInfo();
    }

    private void BuildUi()
    {
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            AutoSize = true,
            AutoScroll = true,
            Padding = new Padding(12),
        };

        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        int row = 0;
        AddRow(table, ref row, nameof(Informace.BasePath), _tbBasePath, CreateBrowseFolderButton(_tbBasePath));
        AddRow(table, ref row, nameof(Informace.SouborStrojeXls), _tbSouborStrojeXls, CreateBrowseFileButton(_tbSouborStrojeXls, "Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Všechny soubory (*.*)|*.*"));
        AddRow(table, ref row, nameof(Informace.SouborStrojeJson), _tbSouborStrojeJson, CreateBrowseFileButton(_tbSouborStrojeJson, "JSON (*.json)|*.json|Všechny soubory (*.*)|*.*"));
        AddRow(table, ref row, nameof(Informace.SouborElektroJson), _tbSouborElektroJson, CreateBrowseFileButton(_tbSouborElektroJson, "JSON (*.json)|*.json|Všechny soubory (*.*)|*.*"));
        AddRow(table, ref row, nameof(Informace.Místnost), _tbMistnost);
        AddRow(table, ref row, nameof(Informace.Projekt), _tbProjekt);
        AddRow(table, ref row, nameof(Informace.Název), _tbNazev);
        AddRow(table, ref row, nameof(Informace.Poznámka), _tbPoznamka);
        AddRow(table, ref row, nameof(Informace.AdresarZdrojDat), _tbData, CreateBrowseFolderButton(_tbData));
        AddRow(table, ref row, nameof(Informace.Datum), _dpDatum);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            WrapContents = false,
        };

        var btnUlozit = new Button { Text = "Uložit", AutoSize = true };
        btnUlozit.Click += (_, _) => SaveAndClose();

        var btnZrusit = new Button { Text = "Zrušit", AutoSize = true, DialogResult = DialogResult.Cancel };
        btnZrusit.Click += (_, _) => Close();

        buttons.Controls.Add(btnUlozit);
        buttons.Controls.Add(btnZrusit);

        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.Controls.Add(buttons, 0, row);
        table.SetColumnSpan(buttons, 3);

        AcceptButton = btnUlozit;
        CancelButton = btnZrusit;

        Controls.Add(table);
        MinimumSize = new System.Drawing.Size(760, 420);
    }

    private static void AddRow(TableLayoutPanel table, ref int row, string propertyName, Control editor, Control? rightButton = null)
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

        table.Controls.Add(label, 0, row);
        table.Controls.Add(editor, 1, row);

        if (rightButton is null)
        {
            table.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, row);
        }
        else
        {
            rightButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            table.Controls.Add(rightButton, 2, row);
        }

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
        var btn = new Button { Text = "...", Width = 36, Height = 26 };
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
        var btn = new Button { Text = "...", Width = 36, Height = 26 };
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
                // Ignorovat chybnou cestu
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
        _tbMistnost.Text = _info.Místnost ?? string.Empty;
        _tbProjekt.Text = _info.Projekt ?? string.Empty;
        _tbNazev.Text = _info.Název ?? string.Empty;
        _tbPoznamka.Text = _info.Poznámka ?? string.Empty;
        _tbData.Text = _info.AdresarZdrojDat ?? string.Empty;
        _dpDatum.Value = _info.Datum == default ? DateTime.Now : _info.Datum;
    }

    private void ApplyToInfo()
    {
        _info.BasePath = _tbBasePath.Text.Trim();
        _info.SouborStrojeXls = _tbSouborStrojeXls.Text.Trim();
        _info.SouborStrojeJson = _tbSouborStrojeJson.Text.Trim();
        _info.SouborElektroJson = _tbSouborElektroJson.Text.Trim();
        _info.Místnost = _tbMistnost.Text.Trim();
        _info.Projekt = _tbProjekt.Text.Trim();
        _info.Název = _tbNazev.Text.Trim();
        _info.Poznámka = _tbPoznamka.Text;
        _info.AdresarZdrojDat = _tbData.Text.Trim();
        _info.Datum = _dpDatum.Value;
    }

    private void SaveAndClose()
    {
        ApplyToInfo();
        _info.Ulozit();
        DialogResult = DialogResult.OK;
        Close();
    }
}

