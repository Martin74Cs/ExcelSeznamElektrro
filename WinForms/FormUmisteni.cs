using Knihovna;
using Knihovna.Shared.Tridy;
using Knihovna.Tridy;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace WinForms
{
    public partial class FormUmisteni : Form
    {
        private readonly List<Zarizeni> _seznamZarizeni;
        private readonly List<string> _seznamHodnot;
        private readonly string? _vychoziVlastnost;
        private readonly string? _filePath;

        private class VlastnostItem
        {
            public string DisplayName { get; }
            public string PropertyName { get; }
            public VlastnostItem(string displayName, string propertyName)
            {
                DisplayName = displayName;
                PropertyName = propertyName;
            }
            public override string ToString() => DisplayName;
        }

        public FormUmisteni(List<Zarizeni> seznamZarizeni, string? vychoziVlastnost = null, string? filePath = null)
        {
            InitializeComponent();
            _seznamZarizeni = seznamZarizeni;
            _seznamHodnot = new List<string>();
            _vychoziVlastnost = vychoziVlastnost;
            _filePath = filePath;
        }

        private void FormUmisteni_Load(object sender, EventArgs e)
        {
            // Inicializace možností pro správu
            comboBoxVlastnost.Items.Add(new VlastnostItem("Rozvaděč (Rozvaděč)", nameof(Zarizeni.RozvadecOznačení)));
            comboBoxVlastnost.Items.Add(new VlastnostItem("Stavební objekt (Objekt)", nameof(Zarizeni.Objekt)));
            comboBoxVlastnost.Items.Add(new VlastnostItem("Provozní soubor (Provozni)", nameof(Zarizeni.Provozni)));
            comboBoxVlastnost.Items.Add(new VlastnostItem("Patro (Patro)", nameof(Zarizeni.Patro)));
            comboBoxVlastnost.Items.Add(new VlastnostItem("Fáze výstavby (Etapa)", nameof(Zarizeni.Etapa)));

            int indexToSelect = 0;
            if (!string.IsNullOrEmpty(_vychoziVlastnost))
            {
                for (int i = 0; i < comboBoxVlastnost.Items.Count; i++)
                {
                    if (comboBoxVlastnost.Items[i] is VlastnostItem item && item.PropertyName == _vychoziVlastnost)
                    {
                        indexToSelect = i;
                        break;
                    }
                }
            }

            if (comboBoxVlastnost.Items.Count > 0)
            {
                comboBoxVlastnost.SelectedIndex = indexToSelect;
            }
        }

        private VlastnostItem? VybranaVlastnost => comboBoxVlastnost.SelectedItem as VlastnostItem;

        private string GetZarizeniValue(Zarizeni zar, string propName)
        {
            if (zar == null || string.IsNullOrEmpty(propName)) return string.Empty;
            PropertyInfo? prop = typeof(Zarizeni).GetProperty(propName);
            return prop?.GetValue(zar) as string ?? string.Empty;
        }

        private void SetZarizeniValue(Zarizeni zar, string propName, string value)
        {
            if (zar == null || string.IsNullOrEmpty(propName)) return;

            if (propName == nameof(Zarizeni.RozvadecOznačení))
            {
                var (text, cislo) = RozdelRozvadec(value);
                zar.Rozvadec = text;
                zar.RozvadecCislo = cislo;

                // Synchronizujeme také u kabelů tohoto zařízení
                if (zar.SeznamKabelu != null)
                {
                    foreach (var k in zar.SeznamKabelu)
                    {
                        k.Rozvadec = text;
                        k.RozvadecCislo = cislo;
                    }
                }
            }
            else
            {
                PropertyInfo? prop = typeof(Zarizeni).GetProperty(propName);
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(zar, value);
                }
            }
        }

        private bool IsZarizeniUnassigned(Zarizeni zar, string propName)
        {
            if (zar == null || string.IsNullOrEmpty(propName)) return true;
            if (propName == nameof(Zarizeni.RozvadecOznačení))
            {
                return string.IsNullOrWhiteSpace(zar.Rozvadec);
            }
            return string.IsNullOrWhiteSpace(GetZarizeniValue(zar, propName));
        }

        private static (string text, string cislo) RozdelRozvadec(string nazev)
        {
            if (string.IsNullOrEmpty(nazev)) return (string.Empty, string.Empty);

            int index = 0;
            while (index < nazev.Length && !char.IsDigit(nazev[index]))
            {
                index++;
            }

            if (index == 0)
            {
                return (string.Empty, nazev);
            }
            if (index == nazev.Length)
            {
                return (nazev, string.Empty);
            }

            return (nazev.Substring(0, index), nazev.Substring(index));
        }

        private void ComboBoxVlastnost_SelectedIndexChanged(object sender, EventArgs e)
        {
            NacistJedinecneHodnoty();
        }

        private void NacistJedinecneHodnoty()
        {
            if (VybranaVlastnost == null) return;

            string propName = VybranaVlastnost.PropertyName;

            // Načtení jedinečných hodnot pro vybranou vlastnost z existujících dat
            var existujici = _seznamZarizeni
                .Select(z => GetZarizeniValue(z, propName))
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct()
                .OrderBy(v => v);

            _seznamHodnot.Clear();
            foreach (var val in existujici)
            {
                _seznamHodnot.Add(val);
            }

            ObnovHodnoty();

            if (listBoxHodnoty.Items.Count > 0)
            {
                listBoxHodnoty.SelectedIndex = 0;
            }
            else
            {
                ObnovZarizeni();
            }
        }

        private void ObnovHodnoty()
        {
            listBoxHodnoty.DataSource = null;
            listBoxHodnoty.DataSource = _seznamHodnot;
        }

        private void ListBoxHodnoty_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObnovZarizeni();
        }

        private void ObnovZarizeni()
        {
            if (VybranaVlastnost == null) return;

            string propName = VybranaVlastnost.PropertyName;
            string vybranaHodnota = listBoxHodnoty.SelectedItem?.ToString() ?? "";

            // Aktualizace textů v groupboxech podle kontextu
            groupBoxPrirazeno.Text = $"Zařízení v hodnotě \"{vybranaHodnota}\"";
            groupBoxNeprirazeno.Text = $"Nepřiřazená zařízení (prázdná hodnota)";

            // Zařízení přiřazená k vybrané hodnotě
            dataGridViewPrirazeno.DataSource = null;
            var prirazena = _seznamZarizeni
                .Where(z => string.Equals(GetZarizeniValue(z, propName), vybranaHodnota, StringComparison.OrdinalIgnoreCase))
                .ToList();
            dataGridViewPrirazeno.DataSource = prirazena;
            NastavSloupceGridu(dataGridViewPrirazeno);

            // Nepřiřazená zařízení
            dataGridViewNeprirazeno.DataSource = null;
            var neprirazena = _seznamZarizeni
                .Where(z => IsZarizeniUnassigned(z, propName))
                .ToList();
            dataGridViewNeprirazeno.DataSource = neprirazena;
            NastavSloupceGridu(dataGridViewNeprirazeno);

            // Statistiky
            int celkem = _seznamZarizeni.Count;
            int prirazenoCelkem = _seznamZarizeni.Count(z => !IsZarizeniUnassigned(z, propName));
            int neprirazenoCelkem = celkem - prirazenoCelkem;

            lblStatistika.Text = $"Statistika:  Celkem: {celkem}  |  Přiřazeno: {prirazenoCelkem}  |  Chybí přiřadit: {neprirazenoCelkem}";
        }

        private void NastavSloupceGridu(DataGridView dgv)
        {
            if (dgv.Columns.Count == 0) return;

            // Zobrazíme jen klíčové sloupce pro přehlednost
            string[] zobrazit = { "Tag", "Popis", "Druh", "Typ", "Objekt", "Provozni", "Patro", "Etapa", "RozvadecOznačení" };
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.Visible = zobrazit.Contains(col.Name);

                if (col.Name == "Tag") col.HeaderText = "Označení (Tag)";
                else if (col.Name == "Popis") col.HeaderText = "Popis";
                else if (col.Name == "Druh") col.HeaderText = "Druh";
                else if (col.Name == "Typ") col.HeaderText = "Typ";
                else if (col.Name == "Objekt") col.HeaderText = "Stavební objekt";
                else if (col.Name == "Provozni") col.HeaderText = "Provozní soubor";
                else if (col.Name == "Patro") col.HeaderText = "Patro";
                else if (col.Name == "Etapa") col.HeaderText = "Etapa/Fáze";
                else if (col.Name == "RozvadecOznačení") col.HeaderText = "Rozvaděč";
            }
        }

        private void BtnPridat_Click(object sender, EventArgs e)
        {
            string nova = txtNovaHodnota.Text.Trim();
            if (string.IsNullOrWhiteSpace(nova))
            {
                MessageBox.Show("Zadejte název nové hodnoty.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_seznamHodnot.Any(h => string.Equals(h, nova, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Tato hodnota již v seznamu existuje.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _seznamHodnot.Add(nova);
            _seznamHodnot.Sort();
            ObnovHodnoty();
            listBoxHodnoty.SelectedItem = nova;
            txtNovaHodnota.Clear();
        }

        private void BtnSloucit_Click(object sender, EventArgs e)
        {
            if (VybranaVlastnost == null) return;
            string propName = VybranaVlastnost.PropertyName;

            string? staraHodnota = listBoxHodnoty.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(staraHodnota))
            {
                MessageBox.Show("Vyberte v seznamu hodnotu, kterou chcete přejmenovat.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string novaHodnota = txtNovaHodnota.Text.Trim();
            if (string.IsNullOrWhiteSpace(novaHodnota))
            {
                MessageBox.Show("Zadejte nový název pro vybranou hodnotu.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.Equals(staraHodnota, novaHodnota, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Zadaný nový název je stejný jako původní.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool jizExistuje = _seznamHodnot.Any(h => string.Equals(h, novaHodnota, StringComparison.OrdinalIgnoreCase));
            string zprava = jizExistuje
                ? $"Cílová hodnota '{novaHodnota}' již existuje. Chcete všechny záznamy z '{staraHodnota}' přesunout a sloučit pod '{novaHodnota}'?"
                : $"Opravdu chcete přejmenovat '{staraHodnota}' na '{novaHodnota}' u všech zařízení?";

            var result = MessageBox.Show(zprava, "Přejmenování / Sloučení", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                // Nahrazení hodnoty u všech zařízení
                foreach (var z in _seznamZarizeni)
                {
                    if (string.Equals(GetZarizeniValue(z, propName), staraHodnota, StringComparison.OrdinalIgnoreCase))
                    {
                        SetZarizeniValue(z, propName, novaHodnota);
                    }
                }

                // Obnovení seznamu hodnot
                if (jizExistuje)
                {
                    _seznamHodnot.Remove(staraHodnota);
                }
                else
                {
                    int index = _seznamHodnot.IndexOf(staraHodnota);
                    if (index >= 0)
                    {
                        _seznamHodnot[index] = novaHodnota;
                    }
                }
                _seznamHodnot.Sort();

                ObnovHodnoty();
                listBoxHodnoty.SelectedItem = novaHodnota;
                txtNovaHodnota.Clear();
                ObnovZarizeni();
            }
        }

        private void BtnPriradit_Click(object sender, EventArgs e)
        {
            if (VybranaVlastnost == null) return;
            string propName = VybranaVlastnost.PropertyName;

            string vybranaHodnota = listBoxHodnoty.SelectedItem?.ToString() ?? "";
            if (string.IsNullOrEmpty(vybranaHodnota))
            {
                MessageBox.Show("Vyberte v levém seznamu cílovou hodnotu.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dataGridViewNeprirazeno.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte v dolní tabulce zařízení, která chcete přiřadit.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (DataGridViewRow row in dataGridViewNeprirazeno.SelectedRows)
            {
                if (row.DataBoundItem is Zarizeni activeZar)
                {
                    SetZarizeniValue(activeZar, propName, vybranaHodnota);
                }
            }

            ObnovZarizeni();
        }

        private void BtnOdebrat_Click(object sender, EventArgs e)
        {
            if (VybranaVlastnost == null) return;
            string propName = VybranaVlastnost.PropertyName;

            if (dataGridViewPrirazeno.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte v horní tabulce zařízení, která chcete odebrat.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string zprava = $"Opravdu chcete odebrat vybraná zařízení?";
            if (MessageBox.Show(zprava, "Potvrzení odebrání", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in dataGridViewPrirazeno.SelectedRows)
                {
                    if (row.DataBoundItem is Zarizeni activeZar)
                    {
                        SetZarizeniValue(activeZar, propName, string.Empty);
                    }
                }

                ObnovZarizeni();
            }
        }

        private void UlozData()
        {
            try
            {
                if (!string.IsNullOrEmpty(_filePath))
                {
                    _seznamZarizeni.SaveJsonList(_filePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Nepodařilo se uložit data do souboru: {ex.Message}", "Chyba ukládání", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnZavrit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            UlozData();
        }
    }
}
