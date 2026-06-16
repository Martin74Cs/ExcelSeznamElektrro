using Knihovna;
using Knihovna.Tridy;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace WinForms
{
    public partial class FormRozvadece : Form
    {
        private readonly List<Zarizeni> _seznamZarizeni;
        private readonly List<string> _seznamRozvadecu;

        public FormRozvadece(List<Zarizeni> seznamZarizeni)
        {
            InitializeComponent();
            _seznamZarizeni = seznamZarizeni;
            _seznamRozvadecu = new List<string>();
        }

        private void FormRozvadece_Load(object sender, EventArgs e)
        {
            // Načtení jedinečných rozvaděčů z existujících dat
            var existujici = _seznamZarizeni
                .Select(z => z.RozvadecOznačení)
                .Where(r => !string.IsNullOrWhiteSpace(r))
                .Distinct()
                .OrderBy(r => r);

            foreach (var r in existujici)
            {
                _seznamRozvadecu.Add(r);
            }

            ObnovRozvadece();
            
            if (listBoxRozvadece.Items.Count > 0)
            {
                listBoxRozvadece.SelectedIndex = 0;
            }

            ObnovZarizeni();
        }

        private void ObnovRozvadece()
        {
            listBoxRozvadece.DataSource = null;
            listBoxRozvadece.DataSource = _seznamRozvadecu;
        }

        private void ListBoxRozvadece_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObnovZarizeni();
        }

        private void ObnovZarizeni()
        {
            string vybranyRozvadec = listBoxRozvadece.SelectedItem?.ToString() ?? "";

            // Zařízení v rozvaděči
            dataGridViewPrirazeno.DataSource = null;
            var prirazena = _seznamZarizeni
                .Where(z => string.Equals(z.RozvadecOznačení, vybranyRozvadec, StringComparison.OrdinalIgnoreCase))
                .ToList();
            dataGridViewPrirazeno.DataSource = prirazena;
            NastavSloupceGridu(dataGridViewPrirazeno);

            // Nepřiřazená zařízení
            dataGridViewNeprirazeno.DataSource = null;
            var neprirazena = _seznamZarizeni
                .Where(z => string.IsNullOrWhiteSpace(z.Rozvadec))
                .ToList();
            dataGridViewNeprirazeno.DataSource = neprirazena;
            NastavSloupceGridu(dataGridViewNeprirazeno);

            // Statistiky
            int celkem = _seznamZarizeni.Count;
            int prirazenoCelkem = _seznamZarizeni.Count(z => !string.IsNullOrWhiteSpace(z.Rozvadec));
            int neprirazenoCelkem = celkem - prirazenoCelkem;

            lblStatistika.Text = $"Statistika:  Celkem: {celkem}  |  Přiřazeno: {prirazenoCelkem}  |  Chybí přiřadit: {neprirazenoCelkem}";
        }

        private void NastavSloupceGridu(DataGridView dgv)
        {
            if (dgv.Columns.Count == 0) return;

            // Zobrazíme jen klíčové sloupce pro přehlednost
            string[] zobrazit = { "Tag", "Popis", "Druh", "Typ", "RozvadecOznačení" };
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.Visible = zobrazit.Contains(col.Name);

                if (col.Name == "Tag") col.HeaderText = "Označení (Tag)";
                else if (col.Name == "Popis") col.HeaderText = "Popis zařízení";
                else if (col.Name == "Druh") col.HeaderText = "Druh";
                else if (col.Name == "Typ") col.HeaderText = "Typ";
                else if (col.Name == "RozvadecOznačení") col.HeaderText = "Rozvaděč";
            }
        }

        private void BtnPridatRozvadec_Click(object sender, EventArgs e)
        {
            string novy = txtNovyRozvadec.Text.Trim();
            if (string.IsNullOrWhiteSpace(novy))
            {
                MessageBox.Show("Zadejte název nového rozvaděče.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_seznamRozvadecu.Contains(novy))
            {
                MessageBox.Show("Tento rozvaděč již v seznamu existuje.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _seznamRozvadecu.Add(novy);
            _seznamRozvadecu.Sort();
            ObnovRozvadece();
            listBoxRozvadece.SelectedItem = novy;
            txtNovyRozvadec.Clear();
        }

        private void BtnPriradit_Click(object sender, EventArgs e)
        {
            string vybranyRozvadec = listBoxRozvadece.SelectedItem?.ToString() ?? "";
            if (string.IsNullOrEmpty(vybranyRozvadec))
            {
                MessageBox.Show("Vyberte rozvaděč, ke kterému chcete zařízení přiřadit.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dataGridViewNeprirazeno.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte v dolní tabulce zařízení, která chcete přiřadit.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var (text, cislo) = RozdelRozvadec(vybranyRozvadec);

            foreach (DataGridViewRow row in dataGridViewNeprirazeno.SelectedRows)
            {
                if (row.DataBoundItem is Zarizeni activeZar)
                {
                    activeZar.Rozvadec = text;
                    activeZar.RozvadecCislo = cislo;

                    // Synchronizujeme rozvaděč také u kabelů tohoto zařízení
                    foreach (var k in activeZar.SeznamKabelu)
                    {
                        k.Rozvadec = text;
                        k.RozvadecCislo = cislo;
                    }
                }
            }

            UlozData();
            ObnovZarizeni();
        }

        private void BtnOdebrat_Click(object sender, EventArgs e)
        {
            if (dataGridViewPrirazeno.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte v horní tabulce zařízení, která chcete z rozvaděče odebrat.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Opravdu chcete odebrat vybraná zařízení z rozvaděče?", "Potvrzení", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in dataGridViewPrirazeno.SelectedRows)
                {
                    if (row.DataBoundItem is Zarizeni activeZar)
                    {
                        activeZar.Rozvadec = string.Empty;
                        activeZar.RozvadecCislo = string.Empty;

                        // Vyčistíme rozvaděč také u kabelů tohoto zařízení
                        foreach (var k in activeZar.SeznamKabelu)
                        {
                            k.Rozvadec = string.Empty;
                            k.RozvadecCislo = string.Empty;
                        }
                    }
                }

                UlozData();
                ObnovZarizeni();
            }
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

        private void UlozData()
        {
            try
            {
                string Cesta = Path.Combine(Informace.Instance.BasePath, "Elektro.Data.Json");
                _seznamZarizeni.SaveJsonList(Cesta);
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
    }
}
