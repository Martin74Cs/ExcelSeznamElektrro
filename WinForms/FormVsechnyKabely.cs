using Knihovna.Shared.Tridy;
using Knihovna.Tridy;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace WinForms
{
    /// <summary>
    /// Formulář pro hromadné zobrazení a správu všech kabelů v projektu.
    /// </summary>
    public partial class FormVsechnyKabely : Form
    {
        private readonly List<Zarizeni> _seznamZarizeni= [];
        private readonly List<KabelRowView> _vsechnyRadky = [];
        private SortableBindingList<KabelRowView> _dataBind = [];

        /// <summary>
        /// Inicializuje novou instanci třídy <see cref="FormVsechnyKabely"/>.
        /// </summary>
        /// <param name="seznamZarizeni">Kompletní seznam zařízení v projektu.</param>
        public FormVsechnyKabely(List<Zarizeni> seznamZarizeni)
        {
            InitializeComponent();
            _seznamZarizeni = seznamZarizeni;
        }

        private void FormVsechnyKabely_Load(object sender, EventArgs e)
        {
            NactiVsechnyKabely();
            dataGridViewKabely.AutoGenerateColumns = true;
            ObnovGrid();
            NastavSloupceGridu();
            ObnovSeznamBezKabelu();
        }

        /// <summary>
        /// Načte všechny kabely ze všech zařízení a obalí je do KabelRowView.
        /// </summary>
        private void NactiVsechnyKabely()
        {
            _vsechnyRadky.Clear();
            foreach (Zarizeni zarizeni in _seznamZarizeni)
            {
                zarizeni.SeznamKabelu ??= [];

                foreach (Trasa trasa in zarizeni.SeznamKabelu)
                {
                    _vsechnyRadky.Add(new KabelRowView(zarizeni, trasa));
                }
            }
        }

        /// <summary>
        /// Nastaví šířky, pořadí a read-only vlastnosti sloupců gridu.
        /// </summary>
        private void NastavSloupceGridu()
        {
            if (dataGridViewKabely.Columns.Count == 0) return;

            // Přejmenování a nastavení sloupců na základě vlastností KabelRowView
            foreach (DataGridViewColumn col in dataGridViewKabely.Columns)
            {
                if (col.Name == nameof(KabelRowView.ZarizeniTag))
                {
                    col.ReadOnly = true;
                    col.DisplayIndex = 0;
                    col.Width = 120;
                }
                else if (col.Name == nameof(KabelRowView.ZarizeniPredmet))
                {
                    col.ReadOnly = true;
                    col.DisplayIndex = 1;
                    col.Width = 150;
                }
                else if (col.Name == nameof(KabelRowView.ZarizeniPopis))
                {
                    col.ReadOnly = true;
                    col.DisplayIndex = 2;
                    col.Width = 200;
                }
                else if (col.Name == nameof(KabelRowView.Oznaceni))
                {
                    col.DisplayIndex = 3;
                    col.Width = 100;
                }
                else if (col.Name == nameof(KabelRowView.Kabel))
                {
                    col.DisplayIndex = 4;
                    col.Width = 120;
                }
                else if (col.Name == nameof(KabelRowView.PocetZil))
                {
                    col.DisplayIndex = 5;
                    col.Width = 80;
                }
                else if (col.Name == nameof(KabelRowView.Prurezmm2))
                {
                    col.DisplayIndex = 6;
                    col.Width = 100;
                }
                else if (col.Name == nameof(KabelRowView.Delka))
                {
                    col.DisplayIndex = 7;
                    col.Width = 80;
                }
                else if (col.Name == nameof(KabelRowView.Popis))
                {
                    col.DisplayIndex = 8;
                    col.Width = 200;
                }
                else if (col.Name == nameof(KabelRowView.RozvadecAll))
                {
                    col.ReadOnly = true;
                    col.DisplayIndex = 9;
                    col.Width = 120;
                }
                else if (col.Name == nameof(KabelRowView.ProudZatizeni))
                {
                    col.ReadOnly = true;
                    col.DisplayIndex = 10;
                    col.Width = 120;
                }
            }

            // Obarvení needitovatelných sloupců na lehce šedou barvu
            foreach (DataGridViewColumn col in dataGridViewKabely.Columns)
            {
                if (col.ReadOnly)
                {
                    col.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(235, 235, 235);
                }
            }
        }

        /// <summary>
        /// Aplikuje vyhledávací filtr a obnoví zdroj dat gridu.
        /// </summary>
        private void ObnovGrid()
        {
            IEnumerable<KabelRowView> filtrovano = _vsechnyRadky;

            string query = textBoxSearch.Text.Trim();
            if (!string.IsNullOrWhiteSpace(query))
            {
                filtrovano = filtrovano.Where(k =>
                    (k.ZarizeniTag != null && k.ZarizeniTag.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (k.ZarizeniPredmet != null && k.ZarizeniPredmet.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (k.Oznaceni != null && k.Oznaceni.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (k.Kabel != null && k.Kabel.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (k.Popis != null && k.Popis.Contains(query, StringComparison.OrdinalIgnoreCase))
                );
            }

            // Na začátku seřadíme podle Tagu zařízení a Označení kabelu
            var serazeno = filtrovano
                .OrderBy(k => k.ZarizeniTag)
                .ThenBy(k => k.Oznaceni)
                .ToList();

            _dataBind = [with(serazeno)];
            dataGridViewKabely.DataSource = _dataBind;

            lblStatistika.Text = $"Počet kabelů celkem: {_vsechnyRadky.Count} (zobrazeno: {serazeno.Count})";
        }

        private void TextBoxSearch_TextChanged(object sender, EventArgs e)
        {
            ObnovGrid();
        }

        private void DataGridViewKabely_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewKabely.CurrentRow?.DataBoundItem is KabelRowView vybrany)
            {
                ZobrazDetailyZarizeni(vybrany.Zarizeni);
            }
            else
            {
                VymazDetailyZarizeni();
            }
        }

        private void ZobrazDetailyZarizeni(Zarizeni zarizeni)
        {
            txtInfoTag.Text = zarizeni.Tag;
            txtInfoPopis.Text = zarizeni.Popis;
            txtInfoPrikon.Text = zarizeni.Prikon;
            txtInfoProud.Text = zarizeni.Proud;
            txtInfoNapeti.Text = zarizeni.Napeti;
            txtInfoMenic.Text = zarizeni.Menic;
            txtInfoDruh.Text = zarizeni.Druh.ToString();
            txtInfoRozvadec.Text = zarizeni.RozvadecOznačení;
        }

        private void VymazDetailyZarizeni()
        {
            txtInfoTag.Text = string.Empty;
            txtInfoPopis.Text = string.Empty;
            txtInfoPrikon.Text = string.Empty;
            txtInfoProud.Text = string.Empty;
            txtInfoNapeti.Text = string.Empty;
            txtInfoMenic.Text = string.Empty;
            txtInfoDruh.Text = string.Empty;
            txtInfoRozvadec.Text = string.Empty;
        }

        /// <summary>
        /// Vygeneruje unikátní označení kabelu pro dané zařízení na základě prefixu.
        /// </summary>
        private static string GenerujUnikantiOznaceni(string znacka, Zarizeni activeZar)
        {
            int maxNum = 0;
            foreach (Trasa k in activeZar.SeznamKabelu)
            {
                if (k.Oznaceni.StartsWith(znacka, StringComparison.OrdinalIgnoreCase))
                {
                    string numPart = k.Oznaceni.Substring(znacka.Length).Trim();
                    if (int.TryParse(numPart, out int num))
                    {
                        if (num > maxNum) maxNum = num;
                    }
                }
            }
            int nextNum = maxNum + 1;
            return $"{znacka} {nextNum:D2}";
        }

        private void BtnAddKabel_Click(object sender, EventArgs e)
        {
            Zarizeni? activeZar;

            // Pokusíme se vzít zařízení z aktuálně vybraného řádku
            if (dataGridViewKabely.CurrentRow?.DataBoundItem is KabelRowView vybrany)
            {
                activeZar = vybrany.Zarizeni;
            }
            else
            {
                // Pokud není nic vybráno, vezmeme první zařízení v seznamu
                activeZar = _seznamZarizeni.FirstOrDefault();
            }

            if (activeZar == null)
            {
                MessageBox.Show("V projektu není definováno žádné zařízení.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Vygenerujeme označení a vytvoříme nový kabel (trasa)
            string noveOznaceni = GenerujUnikantiOznaceni("WL", activeZar);
            var novaTrasa = new Trasa
            {
                Tag = activeZar.Tag,
                Rozvadec = activeZar.Rozvadec,
                RozvadecCislo = activeZar.RozvadecCislo,
                Oznaceni = noveOznaceni,
                Kabel = "CYKY-J",
                PocetZil = string.IsNullOrEmpty(activeZar.Vodice) ? "3" : activeZar.Vodice,
                Prurezmm2 =  "", //string.IsNullOrEmpty(activeZar.PrurezMM2) ? "1.5" : activeZar.PrurezMM2,
                Druh = string.Empty,
                Popis = "Nový kabel",
                Delka = "", // activeZar.Delka > 0 ? activeZar.Delka.ToString("0.##") : "10",
                Patro = activeZar.Patro,
                Predmet = activeZar.Predmet
            };
            novaTrasa.AktualizujKabelData();

            // Přidáme do zařízení
            activeZar.SeznamKabelu.Add(novaTrasa);

            // Přidáme do lokálního seznamu řádků
            var novyRadek = new KabelRowView(activeZar, novaTrasa);
            _vsechnyRadky.Add(novyRadek);

            // Obnovíme grid
            ObnovGrid();

            // Obnovíme seznam bez kabelů
            ObnovSeznamBezKabelu();

            // Označíme nově přidaný řádek v gridu
            foreach (DataGridViewRow row in dataGridViewKabely.Rows)
            {
                if (row.DataBoundItem is KabelRowView k && ReferenceEquals(k.Trasa, novaTrasa))
                {
                    dataGridViewKabely.CurrentCell = row.Cells.Cast<DataGridViewCell>().FirstOrDefault(c => c.Visible);
                    row.Selected = true;
                    break;
                }
            }
        }

        private void BtnCopyKabel_Click(object sender, EventArgs e)
        {
            if (dataGridViewKabely.CurrentRow?.DataBoundItem is not KabelRowView vybrany)
            {
                MessageBox.Show("Není vybrán žádný kabel ke kopírování.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Zarizeni activeZar = vybrany.Zarizeni;
            Trasa staryKabel = vybrany.Trasa;

            // Zjistíme prefix označení (např. "WL" ze "WL 02")
            string prefix = "WL";
            string[] casti = staryKabel.Oznaceni.Split([' '], StringSplitOptions.RemoveEmptyEntries);
            if (casti.Length > 0)
            {
                prefix = casti[0];
            }

            string noveOznaceni = GenerujUnikantiOznaceni(prefix, activeZar);

            // Vytvoříme kopii trasy
            var kopieTrasy = new Trasa
            {
                Tag = activeZar.Tag,
                Rozvadec = staryKabel.Rozvadec,
                RozvadecCislo = staryKabel.RozvadecCislo,
                Oznaceni = noveOznaceni,
                Kabel = staryKabel.Kabel,
                PocetZil = staryKabel.PocetZil,
                Prurezmm2 = staryKabel.Prurezmm2,
                Druh = staryKabel.Druh,
                Popis = staryKabel.Popis + " (kopie)",
                Delka = staryKabel.Delka,
                Patro = staryKabel.Patro,
                Predmet = staryKabel.Predmet,
                OdkudSvorka = staryKabel.OdkudSvorka,
                Mezera = staryKabel.Mezera,
                Svorka = staryKabel.Svorka
            };
            kopieTrasy.AktualizujKabelData();

            // Přidáme do zařízení
            activeZar.SeznamKabelu.Add(kopieTrasy);

            // Přidáme do lokálního seznamu
            var novyRadek = new KabelRowView(activeZar, kopieTrasy);
            _vsechnyRadky.Add(novyRadek);

            // Obnovíme grid
            ObnovGrid();

            // Obnovíme seznam bez kabelů
            ObnovSeznamBezKabelu();

            // Označíme nově přidaný řádek v gridu
            foreach (DataGridViewRow row in dataGridViewKabely.Rows)
            {
                if (row.DataBoundItem is KabelRowView k && ReferenceEquals(k.Trasa, kopieTrasy))
                {
                    dataGridViewKabely.CurrentCell = row.Cells.Cast<DataGridViewCell>().FirstOrDefault(c => c.Visible);
                    row.Selected = true;
                    break;
                }
            }
        }

        private void BtnDeleteKabel_Click(object sender, EventArgs e)
        {
            if (dataGridViewKabely.CurrentRow?.DataBoundItem is not KabelRowView vybrany)
            {
                MessageBox.Show("Není vybrán žádný kabel ke smazání.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string msg = $"Opravdu chcete smazat kabel '{vybrany.Oznaceni}' u zařízení '{vybrany.ZarizeniTag}'?";
            if (MessageBox.Show(msg, "Potvrzení smazání", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                // Odebereme ze zařízení
                vybrany.Zarizeni.SeznamKabelu.Remove(vybrany.Trasa);

                // Odebereme z lokálního seznamu
                _vsechnyRadky.Remove(vybrany);

                // Obnovíme grid
                ObnovGrid();

                // Obnovíme seznam bez kabelů
                ObnovSeznamBezKabelu();
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Aktualizuje rozbalovací seznam zařízení, která dosud nemají žádný kabel.
        /// </summary>
        private void ObnovSeznamBezKabelu()
        {
            var zarizeniBezKabelu = _seznamZarizeni
                .Where(z => z.SeznamKabelu == null || z.SeznamKabelu.Count == 0)
                .OrderBy(z => z.Tag)
                .ToList();

            comboBoxZarizeniBezKabelu.DataSource = null;
            comboBoxZarizeniBezKabelu.DataSource = zarizeniBezKabelu;
            comboBoxZarizeniBezKabelu.DisplayMember = "Tag";

            if (zarizeniBezKabelu.Count > 0)
            {
                comboBoxZarizeniBezKabelu.SelectedIndex = 0;
                lblZarizeniBezKabelu.Enabled = true;
                comboBoxZarizeniBezKabelu.Enabled = true;
                btnAddKabelProZarizeni.Enabled = true;
            }
            else
            {
                lblZarizeniBezKabelu.Enabled = false;
                comboBoxZarizeniBezKabelu.Enabled = false;
                btnAddKabelProZarizeni.Enabled = false;
            }
        }

        private void BtnAddKabelProZarizeni_Click(object sender, EventArgs e)
        {
            if (comboBoxZarizeniBezKabelu.SelectedItem is not Zarizeni activeZar)
            {
                MessageBox.Show("Není vybráno žádné zařízení bez kabelu.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Vygenerujeme označení a vytvoříme nový kabel (trasa)
            string noveOznaceni = GenerujUnikantiOznaceni("WL", activeZar);
            var novaTrasa = new Trasa
            {
                Tag = activeZar.Tag,
                Rozvadec = activeZar.Rozvadec,
                RozvadecCislo = activeZar.RozvadecCislo,
                Oznaceni = noveOznaceni,
                Kabel = "CYKY-J",
                PocetZil = string.IsNullOrEmpty(activeZar.Vodice) ? "3" : activeZar.Vodice,
                Prurezmm2 = "",// string.IsNullOrEmpty(activeZar.PrurezMM2) ? "1.5" : activeZar.PrurezMM2,
                Druh = string.Empty,
                Popis = "Nový kabel",
                Delka = "",//activeZar.Delka > 0 ? activeZar.Delka.ToString("0.##") : "10",
                Patro = activeZar.Patro,
                Predmet = activeZar.Predmet
            };

            activeZar.SeznamKabelu ??= [];
            // Přidáme do zařízení
            activeZar.SeznamKabelu.Add(novaTrasa);

            // Přidáme do lokálního seznamu řádků
            var novyRadek = new KabelRowView(activeZar, novaTrasa);
            _vsechnyRadky.Add(novyRadek);

            // Obnovíme grid
            ObnovGrid();

            // Obnovíme seznam bez kabelů
            ObnovSeznamBezKabelu();

            // Označíme nově přidaný řádek v gridu
            foreach (DataGridViewRow row in dataGridViewKabely.Rows)
            {
                if (row.DataBoundItem is KabelRowView k && ReferenceEquals(k.Trasa, novaTrasa))
                {
                    dataGridViewKabely.CurrentCell = row.Cells.Cast<DataGridViewCell>().FirstOrDefault(c => c.Visible);
                    row.Selected = true;
                    break;
                }
            }
        }
    }
}
