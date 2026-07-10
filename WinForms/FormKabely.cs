using Knihovna;
using Knihovna.Shared.Tridy;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WinForms
{
    public partial class FormKabely: Form {
        private readonly List<Zarizeni> _seznamZarizeni;
        private readonly Zarizeni? _predvoleneZarizeni;
        private bool _nacitani = true;

        /// <summary>
        /// Odkaz na právě upravovaný kabel (null = režim přidávání nového kabelu).
        /// </summary>
        private Trasa? _upravovanyKabel = null;

        private GroupBox? groupBoxProud;
        private Label? lblProudZeme;
        private Label? lblProudStena;
        private Label? lblProudTrubkaStena;
        private Label? lblProudIzolace;
        private Label? lblProudVzduch;

        // Prvky pro výběr kabelu z DB a výpočty
        private GroupBox? groupBoxDbAVypocty;
        private RadioButton? rbCu;
        private RadioButton? rbAl;
        private TextBox? txtHledatDb;
        private ListBox? listBoxDbKabely;
        private TextBox? txtCosPhi;
        private ComboBox? comboBoxUlozeni;

        private Label? lblVypPrikon;
        private Label? lblVypNapeti;
        private Label? lblVypProud;
        private Label? lblVypDelka;
        private Label? lblVypUbytek;
        private Label? lblVypZatizitelnost;
        private Label? lblVypTeplota;
        private Label? lblVypZkrat;
        private Label? lblVypImpedance;
        private Label? lblVypTepelnaOdolnost;

        // Pomocná třída pro zobrazení kabelu v ListBoxu
        private class DbKabelItem
        {
            public Kabel Kabel { get; }
            public DbKabelItem(Kabel kabel) => Kabel = kabel;
            public override string ToString() => $"{Kabel.Name} ({Kabel.Deleni}x{Kabel.SLmm2} mm²)";
        }

        public FormKabely(List<Zarizeni> seznamZarizeni, Zarizeni? predvoleneZarizeni = null) {
            InitializeComponent();
            _seznamZarizeni = seznamZarizeni;
            _predvoleneZarizeni = predvoleneZarizeni;
        }

        private void FormKabely_Load(object sender, EventArgs e) {
            _nacitani = true;

            // Naplnění seznamu značek
            comboBoxZnacka.DataSource = new[] { "WH", "WL", "WS", "WC" };
            comboBoxZnacka.SelectedItem = "WL";

            comboBoxFilterExistElektro.Items.Clear();
            comboBoxFilterExistElektro.Items.AddRange(["Vše", "Ano", "Ne"]);
            comboBoxFilterExistElektro.SelectedIndex = 0;

            // Načtení etap
            var etapy = _seznamZarizeni
                .Select(z => z.Etapa)
                .Where(et => !string.IsNullOrWhiteSpace(et))
                .Distinct()
                .OrderBy(et => et)
                .ToList();

            comboBoxFilterEtapa.Items.Clear();
            comboBoxFilterEtapa.Items.Add("Vše");
            foreach(var etapa in etapy) {
                comboBoxFilterEtapa.Items.Add(etapa);
            }
            comboBoxFilterEtapa.SelectedIndex = 0;

            _nacitani = false;

            // Aplikace filtrů a nastavení výchozího zařízení
            AplikujFiltryZarizeni();

            // Pokud bylo předvoleno konkrétní zařízení, pokusíme se jej vybrat (i přes filtry)
            if(_predvoleneZarizeni != null) {
                // Zrušíme filtry, abychom zajistili, že předvolené zařízení bude viditelné
                _nacitani = true;
                comboBoxFilterExistElektro.SelectedIndex = 0;
                comboBoxFilterEtapa.SelectedIndex = 0;
                _nacitani = false;

                AplikujFiltryZarizeni();
                comboBoxZarizeni.SelectedItem = _predvoleneZarizeni;
            }

            ObnovSeznamKabelu();
            NactiVychoziHodnotyZařízení();
            InicializujPanelProudoveZatizitelnosti();
            InicializujPanelDbAVypocty();
        }

        private void ComboBoxZarizeni_SelectedIndexChanged(object? sender, EventArgs e) {
            // Při změně zařízení zrušíme případný režim úpravy
            ResetFormNaNovaKabel();
            ObnovSeznamKabelu();
            NactiVychoziHodnotyZařízení();
        }

        private void ComboBoxZnacka_SelectedIndexChanged(object? sender, EventArgs e) {
            // Automatický návrh označení jen pokud NEJSME v režimu úpravy
            if(_upravovanyKabel == null) {
                AutoSuggestOznaceni();
            }
        }

        private void ComboBoxFilter_SelectedIndexChanged(object? sender, EventArgs e) {
            if(_nacitani) return;
            AplikujFiltryZarizeni();
        }

        private void AplikujFiltryZarizeni() {
            IEnumerable<Zarizeni> query = _seznamZarizeni;

            // 1. Existence v projektu (IsExist)
            if(CheckBoxIsExist.Checked == true) // Ano
            {
                query = query.Where(z => z.IsExist);
            }
            else // Ne
            {
                query = query.Where(z => !z.IsExist);
            }


            // 2. Elektro blok existuje (IsExistElektro)
            if(comboBoxFilterExistElektro.SelectedIndex == 1) // Ano
            {
                query = query.Where(z => z.IsExistElektro);
            }
            else if(comboBoxFilterExistElektro.SelectedIndex == 2) // Ne
            {
                query = query.Where(z => !z.IsExistElektro);
            }

            // 3. Fáze výstavby (Etapa)
            if(comboBoxFilterEtapa.SelectedItem != null && comboBoxFilterEtapa.SelectedItem.ToString() != "Vše") {
                string vybranaEtapa = comboBoxFilterEtapa.SelectedItem.ToString() ?? string.Empty;
                query = query.Where(z => z.Etapa == vybranaEtapa);
            }

            // 4. Text
            if(!string.IsNullOrEmpty(FiltrText.Text) && FiltrText.Text.ToString() != "Vše") {
                query = query.Where(z => z.Tag.Contains(FiltrText.Text));
            }


            var filtrovanySeznam = query.ToList();

            // Dočasně odpojíme event, abychom zamezili zbytečným aktualizacím
            comboBoxZarizeni.SelectedIndexChanged -= ComboBoxZarizeni_SelectedIndexChanged;
            comboBoxZarizeni.DataSource = null;
            comboBoxZarizeni.DataSource = filtrovanySeznam;
            comboBoxZarizeni.DisplayMember = "Tag";
            comboBoxZarizeni.SelectedIndexChanged += ComboBoxZarizeni_SelectedIndexChanged;

            // Pokus o obnovení původního výběru
            if(comboBoxZarizeni.SelectedItem is Zarizeni predesleVybrane && filtrovanySeznam.Contains(predesleVybrane)) {
                comboBoxZarizeni.SelectedItem = predesleVybrane;
            }
            else if(filtrovanySeznam.Count > 0) {
                comboBoxZarizeni.SelectedIndex = 0;
            }
            else {
                comboBoxZarizeni.SelectedIndex = -1;
            }

            ObnovSeznamKabelu();
            NactiVychoziHodnotyZařízení();
        }

        private void ObnovSeznamKabelu() {
            if(comboBoxZarizeni.SelectedItem is not Zarizeni activeZar) {
                dataGridViewKabely.DataSource = null;
                lblStatistika.Text = "Počet kabelů: 0 (celkem v projektu: 0)";
                return;
            }

            // Bindování seznamu kabelů
            dataGridViewKabely.DataSource = null;
            dataGridViewKabely.DataSource = activeZar.SeznamKabelu;

            // Nastavení sloupců v datagridu
            NastavSloupceGridu();

            // Aktualizace statistik
            int celkemVProjektu = _seznamZarizeni.Sum(z => z.SeznamKabelu.Count);
            lblStatistika.Text = $"Počet kabelů u zařízení: {activeZar.SeznamKabelu.Count}  |  Celkem v projektu: {celkemVProjektu}";

            // Automatický návrh označení jen pokud NEJSME v režimu úpravy
            if(_upravovanyKabel == null) {
                AutoSuggestOznaceni();
            }
            AktualizujZobrazeniProudu();
        }

        private void NastavSloupceGridu() {
            if(dataGridViewKabely.Columns.Count == 0) return;

            // Zobrazíme a popíšeme sloupce
            string[] zobrazit = [ "Oznaceni", "Kabel", "PocetZil", "Prurezmm2", "Delka", "Popis" ];
            foreach(DataGridViewColumn col in dataGridViewKabely.Columns) {
                col.Visible = zobrazit.Contains(col.Name);

                // České popisky záhlaví
                if(col.Name == "Oznaceni") col.HeaderText = "Označení";
                else if(col.Name == "Kabel") col.HeaderText = "Kabel (Typ)";
                else if(col.Name == "PocetZil") col.HeaderText = "Počet žil";
                else if(col.Name == "Prurezmm2") col.HeaderText = "Průřez [mm2]";
                else if(col.Name == "Delka") col.HeaderText = "Délka [m]";
                else if(col.Name == "Popis") col.HeaderText = "Popis / Poznámka";
            }
        }

        private void NactiVychoziHodnotyZařízení() {
            if(comboBoxZarizeni.SelectedItem is not Zarizeni activeZar) {
                // Vyčistíme informační texty o zařízení
                txtInfoPopis.Text = string.Empty;
                txtInfoPrikon.Text = string.Empty;
                txtInfoPrikonStroj.Text = string.Empty;
                txtInfoProud.Text = string.Empty;
                txtInfoNapeti.Text = string.Empty;
                txtInfoMenic.Text = string.Empty;
                txtInfoBalena.Text = string.Empty;
                txtInfoDruh.Text = string.Empty;

                // Vyčistíme pole nového kabelu
                txtPocetZil.Text = string.Empty;
                txtPrurez.Text = string.Empty;
                txtDelka.Text = string.Empty;
                txtPopis.Text = string.Empty;
                return;
            }

            // Načtení informací o vybraném zařízení do horního panelu
            txtInfoPopis.Text = activeZar.Popis;
            txtInfoPrikon.Text = activeZar.Prikon;
            txtInfoPrikonStroj.Text = activeZar.PrikonStroj;
            txtInfoProud.Text = activeZar.Proud;
            txtInfoNapeti.Text = activeZar.Napeti;
            txtInfoMenic.Text = activeZar.Menic;
            txtInfoBalena.Text = activeZar.BalenaJednotka;
            txtInfoDruh.Text = activeZar.Druh.ToString();

            // Předvyplnění hodnot pro nový kabel (pouze pokud NEJSME v režimu úpravy)
            if(_upravovanyKabel == null) {
                txtPocetZil.Text = activeZar.Vodice;
                //txtPrurez.Text = activeZar.PrurezMM2;
                //txtDelka.Text = activeZar.Delka.ToString("0.##");
                txtPopis.Text = string.Empty;
            }

            // Nastavení výchozího účiníku na základě motoru, pokud existuje
            if (txtCosPhi != null)
            {
                double ucinik = 0.85;
                if (activeZar.Motor != null && activeZar.Motor.Ucinik50 > 0)
                {
                    double mUcinik = activeZar.Motor.Ucinik50;
                    // Pokud je hodnota v % (např. 82 nebo 85), převedeme na desetinné číslo
                    if (mUcinik > 1.0) mUcinik /= 100.0;
                    ucinik = mUcinik;
                }
                txtCosPhi.Text = ucinik.ToString("0.##");
            }

            if (lblVypPrikon != null)
            {
                PrepocitejVypocty();
            }
        }

        private static string GenerujUnikantiOznaceni(string znacka, Zarizeni activeZar) {
            int maxNum = 0;
            //foreach(var z in _seznamZarizeni) {
            foreach (var k in activeZar.SeznamKabelu)  {
                if(k.Oznaceni.StartsWith(znacka)) {
                    string numPart = k.Oznaceni.Substring(znacka.Length).Trim();
                    if(int.TryParse(numPart, out int num)) {
                        if(num > maxNum) maxNum = num;
                    }
                }
            }
            //}

            int nextNum = maxNum + 1;
            return $"{znacka} {nextNum:D2}";
        }

        private void AutoSuggestOznaceni() {
            if(comboBoxZarizeni.SelectedItem is not Zarizeni activeZar || comboBoxZnacka.SelectedItem == null)
                return;

            string znacka = comboBoxZnacka.SelectedItem.ToString() ?? "WL";
            txtOznaceni.Text = GenerujUnikantiOznaceni(znacka, activeZar);
        }

        // ===================================================================
        // VALIDACE – kontrola vyplnění polí před přidáním/uložením kabelu
        // ===================================================================

        /// <summary>
        /// Zkontroluje, zda jsou všechna povinná pole korektně vyplněna.
        /// Vrací true pokud je vše v pořádku, jinak zobrazí chybovou hlášku a vrátí false.
        /// </summary>
        private bool ValidateKabel(Zarizeni vybraneZarizeni) {
            // 1. Označení – nesmí být prázdné
            if(string.IsNullOrWhiteSpace(txtOznaceni.Text)) {
                MessageBox.Show("Označení kabelu musí být vyplněno.", "Chyba validace",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtOznaceni.Focus();
                return false;
            }

            string noveOznaceni = txtOznaceni.Text.Trim();

            // 2. Označení – kontrola duplicity (při úpravě ignorujeme stávající kabel)
            //bool duplicita = _seznamZarizeni.Any(z => z.SeznamKabelu.Any(k =>
            //    !ReferenceEquals(k, _upravovanyKabel) &&

            // 2. Označení – kontrola duplicity pouze v rámci jednoho vybraného zařízení
            bool duplicita = vybraneZarizeni.SeznamKabelu.Any(k =>
                !ReferenceEquals(k, _upravovanyKabel) &&
                string.Equals(k.Oznaceni, noveOznaceni, StringComparison.OrdinalIgnoreCase));

            if(duplicita) {
                MessageBox.Show($"Kabel s označením '{noveOznaceni}' již v projektu existuje.",
                    "Chyba validace", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtOznaceni.Focus();
                return false;
            }

            // 3. Typ kabelu – nesmí být prázdný
            if(string.IsNullOrWhiteSpace(txtTyp.Text)) {
                MessageBox.Show("Typ kabelu musí být vyplněn.", "Chyba validace",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTyp.Focus();
                return false;
            }

            // 4. Počet žil – nesmí být prázdný
            if(string.IsNullOrWhiteSpace(txtPocetZil.Text)) {
                MessageBox.Show("Počet žil musí být vyplněn.", "Chyba validace",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPocetZil.Focus();
                return false;
            }

            // 5. Průřez – nesmí být prázdný
            if(string.IsNullOrWhiteSpace(txtPrurez.Text)) {
                MessageBox.Show("Průřez musí být vyplněn.", "Chyba validace",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPrurez.Focus();
                return false;
            }

            // 6. Délka – musí být platné číslo > 0 (ošetření čárek i teček)
            string delkaText = txtDelka.Text.Trim().Replace(',', '.');
            if(!double.TryParse(delkaText, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double delkaValue) || delkaValue <= 0) {
                MessageBox.Show("Délka musí být platné číslo větší než 0.", "Chyba validace",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDelka.Focus();
                return false;
            }

            return true;
        }

        // ===================================================================
        // PŘIDÁNÍ / ULOŽENÍ kabelu (sdílené tlačítko btnPridat)
        // ===================================================================

        private void BtnPridat_Click(object sender, EventArgs e) {
            if(comboBoxZarizeni.SelectedItem is not Zarizeni activeZar) {
                MessageBox.Show("Není vybráno žádné zařízení.", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Spustíme validaci
            if(!ValidateKabel(activeZar))
                return;

            string noveOznaceni = txtOznaceni.Text.Trim();

            if(_upravovanyKabel != null) {
                // ────────────────────── REŽIM ÚPRAVY ──────────────────────
                _upravovanyKabel.Oznaceni = noveOznaceni;
                _upravovanyKabel.Kabel = txtTyp.Text.Trim();
                _upravovanyKabel.PocetZil = txtPocetZil.Text.Trim();
                _upravovanyKabel.Prurezmm2 = txtPrurez.Text.Trim();
                _upravovanyKabel.Delka = txtDelka.Text.Trim();
                _upravovanyKabel.Svorka = txtSvorka.Text.Trim();
                _upravovanyKabel.Popis = txtPopis.Text.Trim();
                _upravovanyKabel.AktualizujKabelData();

                ResetFormNaNovaKabel();
                ObnovSeznamKabelu();
            }
            else {
                // ────────────────────── REŽIM PŘIDÁNÍ ──────────────────────
                var trasa = new Trasa {
                    Tag = activeZar.Tag,
                    Rozvadec = activeZar.Rozvadec,
                    RozvadecCislo = activeZar.RozvadecCislo,
                    Oznaceni = noveOznaceni,
                    Kabel = txtTyp.Text.Trim(),
                    PocetZil = txtPocetZil.Text.Trim(),
                    Prurezmm2 = txtPrurez.Text.Trim(),
                    Druh = string.Empty, // Nahrazeno popisem, v databázi necháme prázdný
                    Popis = txtPopis.Text.Trim(),
                    Delka = txtDelka.Text.Trim(),
                    Patro = activeZar.Patro,
                    Predmet = activeZar.Predmet
                };
                trasa.AktualizujKabelData();

                activeZar.SeznamKabelu.Add(trasa);

                ObnovSeznamKabelu();
            }
        }

        private void BtnSmazat_Click(object sender, EventArgs e) {
            if(comboBoxZarizeni.SelectedItem is not Zarizeni activeZar || dataGridViewKabely.CurrentRow == null)
                return;

            if(dataGridViewKabely.CurrentRow.DataBoundItem is Trasa vybranyKabel) {
                if(MessageBox.Show($"Opravdu chcete smazat kabel '{vybranyKabel.Oznaceni}'?", "Potvrzení", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    activeZar.SeznamKabelu.Remove(vybranyKabel);
                    ResetFormNaNovaKabel();
                    ObnovSeznamKabelu();
                }
            }
        }

        // ===================================================================
        // DVOJKLIK na řádek v tabulce kabelů – přepnutí do režimu editace
        // ===================================================================

        private void DataGridViewKabely_CellDoubleClick(object sender, DataGridViewCellEventArgs e) {
            // Ignorujeme kliknutí na záhlaví
            if(e.RowIndex < 0) return;

            if(dataGridViewKabely.Rows[e.RowIndex].DataBoundItem is Trasa vybranyKabel) {
                // Uložíme referenci na upravovaný kabel
                _upravovanyKabel = vybranyKabel;

                // Změna nadpisu panelu a tlačítek
                groupBoxPridat.Text = "Úprava kabelu";
                btnPridat.Text = "Uložit";
                btnStorno.Visible = true;

                // Naplnění polí hodnotami vybraného kabelu
                txtOznaceni.Text = vybranyKabel.Oznaceni;
                txtTyp.Text = vybranyKabel.Kabel;
                txtPocetZil.Text = vybranyKabel.PocetZil;
                txtPrurez.Text = vybranyKabel.Prurezmm2;
                txtDelka.Text = vybranyKabel.Delka;
                txtPopis.Text = vybranyKabel.Popis;

                // Zvýrazníme pole označení pro uživatele
                txtOznaceni.Focus();
                txtOznaceni.SelectAll();
            }
        }

        // ===================================================================
        // STORNO – zrušení režimu úpravy a návrat k přidávání
        // ===================================================================

        private void BtnStorno_Click(object sender, EventArgs e) {
            ResetFormNaNovaKabel();
        }

        /// <summary>
        /// Resetuje editační panel do výchozího stavu "Nový kabel".
        /// </summary>
        private void ResetFormNaNovaKabel() {
            _upravovanyKabel = null;

            // Obnovení nadpisu a tlačítek
            groupBoxPridat.Text = "Nový kabel";
            btnPridat.Text = "Přidat kabel";
            btnStorno.Visible = false;

            // Obnovení výchozích hodnot z aktuálně vybraného zařízení
            if(comboBoxZarizeni.SelectedItem is Zarizeni activeZar) {
                txtPocetZil.Text = activeZar.Vodice;
                txtPrurez.Text = "";// activeZar.PrurezMM2;
                txtDelka.Text = ""; // activeZar.Delka.ToString("0.##");
                txtPopis.Text = string.Empty;
                txtTyp.Text = "JZ-500";
                AutoSuggestOznaceni();
            }
            else {
                txtOznaceni.Text = string.Empty;
                txtTyp.Text = string.Empty;
                txtPocetZil.Text = string.Empty;
                txtPrurez.Text = string.Empty;
                txtDelka.Text = string.Empty;
                txtPopis.Text = string.Empty;
            }
        }

        // ===================================================================
        // RYCHLÉ PŘIDÁNÍ kabelu s nastavitelným prefixem značení
        // ===================================================================

        private void PridejRychlyKabel(string znacka, string typKabelu, string pocetZil, string prurez, string popisKabelu) {
            if(comboBoxZarizeni.SelectedItem is not Zarizeni activeZar) {
                MessageBox.Show("Není vybráno žádné zařízení.", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string noveOznaceni = GenerujUnikantiOznaceni(znacka, activeZar);

            var trasa = new Trasa {
                Tag = activeZar.Tag,
                Rozvadec = activeZar.Rozvadec,
                RozvadecCislo = activeZar.RozvadecCislo,
                Oznaceni = noveOznaceni,
                Kabel = typKabelu,
                PocetZil = pocetZil,
                Prurezmm2 = prurez,
                Druh = string.Empty,
                Popis = popisKabelu,
                Delka = "", // activeZar.Delka.ToString("0.##"),
                Patro = activeZar.Patro,
                Predmet = activeZar.Predmet
            };
            trasa.AktualizujKabelData();

            activeZar.SeznamKabelu.Add(trasa);

            ObnovSeznamKabelu();
        }

        private void BtnRychlyPTC_Click(object sender, EventArgs e) {
            // PTC kabel má typ CYKY-O 2x1.5, 2 žíly, průřez 1.5, popis "PTC čidlo"
            string prefix = string.IsNullOrWhiteSpace(txtPrefixPTC.Text) ? "WS" : txtPrefixPTC.Text.Trim();
            PridejRychlyKabel(prefix, "F-CY-OZ", "2", "1.5", "PTC");
        }

        private void BtnRychlyOvladani5_Click(object sender, EventArgs e) {
            // Ovládací skříň 5 vodičů - typ CYKY-J 5x1.5, 5 žil, průřez 1.5, popis "Ovládací skříňka"
            string prefix = string.IsNullOrWhiteSpace(txtPrefixOvladani.Text) ? "WS" : txtPrefixOvladani.Text.Trim();
            PridejRychlyKabel(prefix, "CYKY-J", "5", "2.5", "Ovládací skříňka");
        }

        private void BtnRychlyOvladani7_Click(object sender, EventArgs e) {
            // Ovládací skříň 7 vodičů - typ CYKY-O 7x1.5, 7 žil, průřez 1.5, popis "Ovládací skříňka"
            string prefix = string.IsNullOrWhiteSpace(txtPrefixOvladani.Text) ? "WS" : txtPrefixOvladani.Text.Trim();
            PridejRychlyKabel(prefix, "JZ-600-Y-CY", "7G", "1.5", "MS");
        }

        private void BtnRychlyOvladani12_Click(object sender, EventArgs e) {
            // Ovládací skříň 12 vodičů - typ CYKY-O 12x1.5, 12 žil, průřez 1.5, popis "Ovládací skříňka"
            string prefix = string.IsNullOrWhiteSpace(txtPrefixOvladani.Text) ? "WS" : txtPrefixOvladani.Text.Trim();
            PridejRychlyKabel(prefix, "CYKY-O", "12", "2.5", "Ovládací skříňka");
        }

        private void BtnRychlyUTP_Click(object sender, EventArgs e) {
            // UTP datová komunikace - typ UTP Cat6, 8 žil, průřez 0.5 (AWG24), popis "UTP datová komunikace"
            string prefix = string.IsNullOrWhiteSpace(txtPrefixUTP.Text) ? "WD" : txtPrefixUTP.Text.Trim();
            PridejRychlyKabel(prefix, "UTP Cat6", "8", "0.5", "UTP datová komunikace");
        }

        private void BtnRychlyBinarni_Click(object sender, EventArgs e) {
            // Binární diskrétní signály - typ CYKY-O 4x1.5, 4 žíly, průřez 1.5, popis "Binární komunikace (diskrétní signály)"
            string prefix = string.IsNullOrWhiteSpace(txtPrefixBinarni.Text) ? "XB" : txtPrefixBinarni.Text.Trim();
            PridejRychlyKabel(prefix, "CYKY-O", "7", "1.5", "Binární komunikace (diskrétní signály)");
        }

        private void BtnRychlyBlokovani_Click(object sender, EventArgs e) {
            // Blokování - typ CYKY-O 3x1.5, 3 žíly, průřez 1.5, popis "Blokování"
            string prefix = string.IsNullOrWhiteSpace(txtPrefixBlokovani.Text) ? "WB" : txtPrefixBlokovani.Text.Trim();
            PridejRychlyKabel(prefix, "CYKY-O", "3", "1.5", "Blokování");
        }


        private void BtnZavrit_Click(object sender, EventArgs e) {
            this.Close();
        }

        private void GroupBoxPridat_Enter(object sender, EventArgs e) {

        }

        private void CheckBoxIsExiste_CheckedChanged(object sender, EventArgs e) {
            if(_nacitani) return;
            AplikujFiltryZarizeni();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e) {
            if(_nacitani) return;
            AplikujFiltryZarizeni();
        }

        private void Button1_Click(object sender, EventArgs e) {
            // FM
            string prefix = string.IsNullOrWhiteSpace(txtPrefixPower.Text) ? "WL" : txtPrefixPower.Text.Trim();
            PridejRychlyKabel(prefix, "JZ-600-Y-CY", "4G", "2.5", "M");
        }

        private void DataGridViewKabely_CellContentClick(object sender, DataGridViewCellEventArgs e) {
            if (e.RowIndex < 0) {
                Console.WriteLine("Špatný index");
                DialogResult = DialogResult.Cancel;
            }
                return;
        }

        private void InicializujPanelProudoveZatizitelnosti()
        {
            groupBoxProud = new GroupBox
            {
                Text = "Proudová zatížitelnost vybraného kabelu [A]",
                Location = new Point(12, 690),
                Size = new Size(796, 100),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            var fRegular = new Font("Segoe UI", 9.5F, FontStyle.Regular);

            var lblZemeTitle = new Label { Text = "V zemi (IzE):", Location = new Point(15, 30), Size = new Size(130, 25), Font = fRegular };
            lblProudZeme = new Label { Text = "-", Location = new Point(15, 60), Size = new Size(130, 25), ForeColor = Color.Blue };

            var lblStenaTitle = new Label { Text = "Na stěně (IzC):", Location = new Point(160, 30), Size = new Size(130, 25), Font = fRegular };
            lblProudStena = new Label { Text = "-", Location = new Point(160, 60), Size = new Size(130, 25), ForeColor = Color.Blue };

            var lblTrubkaTitle = new Label { Text = "V trubce (IzB):", Location = new Point(310, 30), Size = new Size(130, 25), Font = fRegular };
            lblProudTrubkaStena = new Label { Text = "-", Location = new Point(310, 60), Size = new Size(130, 25), ForeColor = Color.Blue };

            var lblIzolaceTitle = new Label { Text = "V izolaci (IzA):", Location = new Point(460, 30), Size = new Size(130, 25), Font = fRegular };
            lblProudIzolace = new Label { Text = "-", Location = new Point(460, 60), Size = new Size(130, 25), ForeColor = Color.Blue };

            var lblVzduchTitle = new Label { Text = "Ve vzduchu (IzG vodorovně / svisle):", Location = new Point(610, 30), Size = new Size(175, 25), Font = fRegular };
            lblProudVzduch = new Label { Text = "- / -", Location = new Point(610, 60), Size = new Size(175, 25), ForeColor = Color.Blue };

            groupBoxProud.Controls.AddRange([
                lblZemeTitle, lblProudZeme,
                lblStenaTitle, lblProudStena,
                lblTrubkaTitle, lblProudTrubkaStena,
                lblIzolaceTitle, lblProudIzolace,
                lblVzduchTitle, lblProudVzduch
            ]);

            this.Controls.Add(groupBoxProud);

            // Napojení na změnu výběru v DataGridView
            dataGridViewKabely.SelectionChanged += DataGridViewKabely_SelectionChanged;

            // Prvotní zobrazení
            AktualizujZobrazeniProudu();
        }

        private void DataGridViewKabely_SelectionChanged(object? sender, EventArgs e)
        {
            AktualizujZobrazeniProudu();
        }

        private void AktualizujZobrazeniProudu()
        {
            if (lblProudZeme == null || lblProudStena == null || lblProudTrubkaStena == null || lblProudIzolace == null || lblProudVzduch == null) return;

            if (dataGridViewKabely.CurrentRow?.DataBoundItem is Trasa vybranyKabel && vybranyKabel.KabelData != null)
            {
                var kd = vybranyKabel.KabelData;
                lblProudZeme.Text = kd.IzAE > 0 ? $"{kd.IzAE} A" : "-";
                lblProudStena.Text = kd.IzAC > 0 ? $"{kd.IzAC} A" : "-";
                lblProudTrubkaStena.Text = kd.IzAB > 0 ? $"{kd.IzAB} A" : "-";
                lblProudIzolace.Text = kd.IzAA > 0 ? $"{kd.IzAA} A" : "-";

                string vodorovne = kd.IzAGvod > 0 ? $"{kd.IzAGvod} A" : "-";
                string svisle = kd.IzAGsvis > 0 ? $"{kd.IzAGsvis} A" : "-";
                lblProudVzduch.Text = $"{vodorovne} / {svisle}";
            }
            else
            {
                lblProudZeme.Text = "-";
                lblProudStena.Text = "-";
                lblProudTrubkaStena.Text = "-";
                lblProudIzolace.Text = "-";
                lblProudVzduch.Text = "- / -";
            }
        }

        private void InicializujPanelDbAVypocty()
        {
            this.AutoSize = false;
            this.ClientSize = new Size(1600, 855);

            groupBoxDbAVypocty = new GroupBox
            {
                Text = "Výběr z databáze a výpočty",
                Location = new Point(1160, 12),
                Size = new Size(420, 831),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            var fRegular = new Font("Segoe UI", 9.5F, FontStyle.Regular);
            var fBold = new Font("Segoe UI", 9.5F, FontStyle.Bold);

            // 1. Výběr Cu / Al
            var lblMaterial = new Label { Text = "Materiál kabelu:", Location = new Point(15, 30), Size = new Size(120, 25), Font = fRegular };
            rbCu = new RadioButton { Text = "Měď (Cu)", Location = new Point(140, 30), Size = new Size(100, 25), Font = fRegular, Checked = true };
            rbAl = new RadioButton { Text = "Hliník (Al)", Location = new Point(250, 30), Size = new Size(110, 25), Font = fRegular };

            rbCu.CheckedChanged += (s, e) => { if (rbCu.Checked) { txtTyp.Text = "CYKY"; AutoSuggestOznaceni(); } AktualizujSeznamDbKabelu(); };
            rbAl.CheckedChanged += (s, e) => { if (rbAl.Checked) { txtTyp.Text = "AYKY"; AutoSuggestOznaceni(); } AktualizujSeznamDbKabelu(); };

            // 2. Vyhledávání a účiník
            var lblHledat = new Label { Text = "Hledat kabel (název/průřez):", Location = new Point(15, 65), Size = new Size(200, 25), Font = fRegular };
            txtHledatDb = new TextBox { Location = new Point(15, 90), Size = new Size(220, 27), Font = fRegular };
            txtHledatDb.TextChanged += (s, e) => AktualizujSeznamDbKabelu();

            var lblCosPhi = new Label { Text = "Účiník (cos φ):", Location = new Point(250, 65), Size = new Size(140, 25), Font = fRegular };
            txtCosPhi = new TextBox { Text = "0,85", Location = new Point(250, 90), Size = new Size(140, 27), Font = fRegular };
            txtCosPhi.TextChanged += (s, e) => PrepocitejVypocty();

            // 2.5 Způsob uložení kabelu
            var lblUlozeni = new Label { Text = "Způsob uložení kabelu:", Location = new Point(15, 125), Size = new Size(380, 20), Font = fRegular };
            comboBoxUlozeni = new ComboBox 
            { 
                Location = new Point(15, 147), 
                Size = new Size(385, 27), 
                Font = fRegular, 
                DropDownStyle = ComboBoxStyle.DropDownList 
            };
            comboBoxUlozeni.Items.AddRange(new object[] 
            { 
                "Na stěně (C)", 
                "Ve vzduchu (E/G)", 
                "V trubce na stěně (B)", 
                "V izolaci (A)", 
                "V zemi (D)" 
            });
            comboBoxUlozeni.SelectedIndex = 0;
            comboBoxUlozeni.SelectedIndexChanged += (s, e) => PrepocitejVypocty();

            // 3. ListBox pro dostupné kabely
            listBoxDbKabely = new ListBox { Location = new Point(15, 185), Size = new Size(385, 160), Font = fRegular };
            listBoxDbKabely.SelectedIndexChanged += ListBoxDbKabely_SelectedIndexChanged;

            // 4. Sekce Výpočty (Záhlaví)
            var lblSekceVypocty = new Label 
            { 
                Text = "Průběžné výsledky výpočtů", 
                Location = new Point(15, 360), 
                Size = new Size(380, 25), 
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.DarkSlateBlue
            };

            // Vytvoříme pomocné labely pro zobrazení
            int startY = 395;
            int stepY = 37;
            
            var lblPrikonTitle = new Label { Text = "Příkon spotřebiče:", Location = new Point(15, startY), Size = new Size(180, 25), Font = fRegular };
            lblVypPrikon = new Label { Text = "-", Location = new Point(200, startY), Size = new Size(210, 25), Font = fBold };

            var lblNapetiTitle = new Label { Text = "Napětí soustavy:", Location = new Point(15, startY + stepY), Size = new Size(180, 25), Font = fRegular };
            lblVypNapeti = new Label { Text = "-", Location = new Point(200, startY + stepY), Size = new Size(210, 25), Font = fBold };

            var lblProudTitle = new Label { Text = "Výpočtový proud (zatížení):", Location = new Point(15, startY + stepY * 2), Size = new Size(180, 25), Font = fRegular };
            lblVypProud = new Label { Text = "-", Location = new Point(200, startY + stepY * 2), Size = new Size(210, 25), Font = fBold, ForeColor = Color.DarkBlue };

            var lblDelkaTitle = new Label { Text = "Délka kabelu:", Location = new Point(15, startY + stepY * 3), Size = new Size(180, 25), Font = fRegular };
            lblVypDelka = new Label { Text = "-", Location = new Point(200, startY + stepY * 3), Size = new Size(210, 25), Font = fBold };

            var lblUbytekTitle = new Label { Text = "Úbytek napětí ΔU:", Location = new Point(15, startY + stepY * 4), Size = new Size(180, 25), Font = fRegular };
            lblVypUbytek = new Label { Text = "-", Location = new Point(200, startY + stepY * 4), Size = new Size(210, 25), Font = fBold, ForeColor = Color.DarkGreen };

            var lblZatizTitle = new Label { Text = "Dovolené zatížení Iz:", Location = new Point(15, startY + stepY * 5), Size = new Size(180, 25), Font = fRegular };
            lblVypZatizitelnost = new Label { Text = "-", Location = new Point(200, startY + stepY * 5), Size = new Size(210, 25), Font = fBold };

            var lblTeplotaTitle = new Label { Text = "Teplota vodiče při zatížení:", Location = new Point(15, startY + stepY * 6), Size = new Size(180, 25), Font = fRegular };
            lblVypTeplota = new Label { Text = "-", Location = new Point(200, startY + stepY * 6), Size = new Size(210, 25), Font = fBold };

            var lblZkratTitle = new Label { Text = "Zkrat sítě Ik (3f/1f):", Location = new Point(15, startY + stepY * 7), Size = new Size(180, 25), Font = fRegular };
            lblVypZkrat = new Label { Text = "-", Location = new Point(200, startY + stepY * 7), Size = new Size(210, 25), Font = fBold };

            var lblTepelnaTitle = new Label { Text = "Tepelná odolnost Ith:", Location = new Point(15, startY + stepY * 8), Size = new Size(180, 25), Font = fRegular };
            lblVypTepelnaOdolnost = new Label { Text = "-", Location = new Point(200, startY + stepY * 8), Size = new Size(210, 25), Font = fBold };

            var lblImpedanceTitle = new Label { Text = "Impedance smyčky Z:", Location = new Point(15, startY + stepY * 9), Size = new Size(180, 25), Font = fRegular };
            lblVypImpedance = new Label { Text = "-", Location = new Point(200, startY + stepY * 9), Size = new Size(210, 25), Font = fBold, ForeColor = Color.Maroon };

            groupBoxDbAVypocty.Controls.AddRange(new Control[]
            {
                lblMaterial, rbCu, rbAl,
                lblHledat, txtHledatDb,
                lblCosPhi, txtCosPhi,
                lblUlozeni, comboBoxUlozeni,
                listBoxDbKabely,
                lblSekceVypocty,
                lblPrikonTitle, lblVypPrikon,
                lblNapetiTitle, lblVypNapeti,
                lblProudTitle, lblVypProud,
                lblDelkaTitle, lblVypDelka,
                lblUbytekTitle, lblVypUbytek,
                lblZatizTitle, lblVypZatizitelnost,
                lblTeplotaTitle, lblVypTeplota,
                lblZkratTitle, lblVypZkrat,
                lblTepelnaTitle, lblVypTepelnaOdolnost,
                lblImpedanceTitle, lblVypImpedance
            });

            this.Controls.Add(groupBoxDbAVypocty);

            // Zajištění automatického přepočítávání při editaci stávajících polí
            txtDelka.TextChanged += (s, e) => PrepocitejVypocty();
            txtPocetZil.TextChanged += (s, e) => PrepocitejVypocty();
            txtPrurez.TextChanged += (s, e) => PrepocitejVypocty();
            txtTyp.TextChanged += (s, e) => PrepocitejVypocty();

            // Nastavení výchozích hodnot pro aktuální zařízení
            NactiVychoziHodnotyZařízení();

            // Načtení databáze kabelů
            AktualizujSeznamDbKabelu();
        }

        private void AktualizujSeznamDbKabelu()
        {
            if (txtHledatDb == null || listBoxDbKabely == null || rbAl == null) return;

            string vyhledavanyText = txtHledatDb.Text.Trim();
            var kabely = rbAl.Checked ? KabelDatabaze.AlKabely : KabelDatabaze.CuKabely;

            IEnumerable<Kabel> filtrovane = kabely;
            if (!string.IsNullOrEmpty(vyhledavanyText))
            {
                filtrovane = filtrovane.Where(k => 
                    k.Name.Contains(vyhledavanyText, StringComparison.OrdinalIgnoreCase) ||
                    k.Označení.Contains(vyhledavanyText, StringComparison.OrdinalIgnoreCase) ||
                    k.Deleni.Contains(vyhledavanyText, StringComparison.OrdinalIgnoreCase) ||
                    k.SLmm2.ToString().Contains(vyhledavanyText, StringComparison.OrdinalIgnoreCase)
                );
            }

            var items = filtrovane.Select(k => new DbKabelItem(k)).ToList();
            
            listBoxDbKabely.SelectedIndexChanged -= ListBoxDbKabely_SelectedIndexChanged;
            listBoxDbKabely.DataSource = null;
            listBoxDbKabely.DataSource = items;
            listBoxDbKabely.SelectedIndexChanged += ListBoxDbKabely_SelectedIndexChanged;

            if (items.Count > 0)
            {
                listBoxDbKabely.SelectedIndex = 0;
            }
            else
            {
                PrepocitejVypocty();
            }
        }

        private void ListBoxDbKabely_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (listBoxDbKabely == null || listBoxDbKabely.SelectedItem is not DbKabelItem dbItem) return;

            var kabel = dbItem.Kabel;
            
            // Přepíšeme hodnoty do stávajících textboxů
            txtTyp.Text = kabel.Name;
            txtPocetZil.Text = kabel.Deleni;
            txtPrurez.Text = kabel.SLmm2.ToString("0.##");
            
            PrepocitejVypocty();
        }

        private void PrepocitejVypocty()
        {
            if (lblVypPrikon == null || lblVypNapeti == null || lblVypProud == null || lblVypDelka == null ||
                lblVypUbytek == null || lblVypZatizitelnost == null || lblVypTeplota == null || 
                lblVypZkrat == null || lblVypTepelnaOdolnost == null || lblVypImpedance == null ||
                comboBoxZarizeni.SelectedItem is not Zarizeni activeZar)
            {
                VynulujVypocty();
                return;
            }

            // 1. Získání příkonu [kW] a jmenovitého napětí [V]
            double prikon = 0;
            string prikonStr = activeZar.Prikon;
            if (string.IsNullOrEmpty(prikonStr)) prikonStr = activeZar.PrikonStroj;
            
            prikonStr = prikonStr.Replace(" ", "").Replace(',', '.');
            double.TryParse(prikonStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out prikon);

            double napeti = 400;
            string napetiStr = activeZar.Napeti;
            if (!string.IsNullOrEmpty(napetiStr))
            {
                napetiStr = napetiStr.Replace(" ", "").Replace(',', '.');
                double.TryParse(napetiStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out napeti);
            }
            if (napeti <= 0) napeti = 400;

            // 2. Délka kabelu [m]
            double delka = 0;
            string delkaStr = txtDelka.Text;
            if (!string.IsNullOrEmpty(delkaStr))
            {
                delkaStr = delkaStr.Replace(" ", "").Replace(',', '.');
                double.TryParse(delkaStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out delka);
            }

            // 3. Účiník cos(phi)
            double cosPhi = 0.85;
            if (txtCosPhi != null && !string.IsNullOrEmpty(txtCosPhi.Text))
            {
                string cosPhiStr = txtCosPhi.Text.Replace(" ", "").Replace(',', '.');
                double.TryParse(cosPhiStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out cosPhi);
            }
            if (cosPhi <= 0 || cosPhi > 1.0) cosPhi = 0.85;

            // 4. Výpočtový proud (zatížení)
            double proudZat = Extension.VypoctiProudZatizeni(prikon, napeti, cosPhi);
            
            // Fallback na proud z parametrů, pokud je příkon nulový
            if (proudZat <= 0)
            {
                string proudStr = activeZar.Proud.Replace(" ", "").Replace(',', '.');
                double.TryParse(proudStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out proudZat);
            }

            // Najdeme kabel
            Kabel? kabel = null;
            if (listBoxDbKabely != null && listBoxDbKabely.SelectedItem is DbKabelItem dbItem)
            {
                kabel = dbItem.Kabel;
            }
            if (kabel == null)
            {
                kabel = KabelDatabaze.NajdiKabel(txtTyp.Text, txtPocetZil.Text, txtPrurez.Text);
            }

            // Zobrazení základních hodnot
            lblVypPrikon.Text = $"{prikon:0.##} kW";
            lblVypNapeti.Text = $"{napeti:0} V";
            lblVypProud.Text = proudZat > 0 ? $"{proudZat:0.##} A" : "-";
            lblVypDelka.Text = $"{delka:0.##} m";

            if (kabel == null)
            {
                lblVypUbytek.Text = "-";
                lblVypZatizitelnost.Text = "-";
                lblVypTeplota.Text = "-";
                lblVypZkrat.Text = "-";
                lblVypTepelnaOdolnost.Text = "-";
                lblVypImpedance.Text = "-";
                return;
            }

            // 4.5 Dovolené zatížení podle způsobu uložení
            double iz = kabel.IzAC;
            string ulozeniZkratka = "C";
            if (comboBoxUlozeni != null)
            {
                switch (comboBoxUlozeni.SelectedIndex)
                {
                    case 0: // Na stěně (C)
                        iz = kabel.IzAC;
                        ulozeniZkratka = "C";
                        break;
                    case 1: // Ve vzduchu (E/G)
                        iz = kabel.IzAGvod > 0 ? kabel.IzAGvod : (kabel.IzAFtroj > 0 ? kabel.IzAFtroj : kabel.IzAC);
                        ulozeniZkratka = "E/G";
                        break;
                    case 2: // V trubce na stěně (B)
                        iz = kabel.IzAB > 0 ? kabel.IzAB : kabel.IzAC;
                        ulozeniZkratka = "B";
                        break;
                    case 3: // V izolaci (A)
                        iz = kabel.IzAA > 0 ? kabel.IzAA : kabel.IzAC;
                        ulozeniZkratka = "A";
                        break;
                    case 4: // V zemi (D)
                        iz = kabel.IzAE > 0 ? kabel.IzAE : kabel.IzAC;
                        ulozeniZkratka = "D";
                        break;
                }
            }
            if (iz <= 0) iz = kabel.IzAC > 0 ? kabel.IzAC : 1.0;
            lblVypZatizitelnost.Text = $"{iz:0} A (metoda {ulozeniZkratka})";

            // 5. Úbytek napětí [%]
            if (proudZat > 0 && delka > 0)
            {
                double ubytekV = kabel.VypoctiUbytekNapetiV(proudZat, delka, napeti, cosPhi);
                double ubytekPct = kabel.VypoctiUbytekNapetiProcenta(proudZat, delka, napeti, cosPhi);
                lblVypUbytek.Text = $"{ubytekPct:0.##} % ({ubytekV:0.##} V)";
            }
            else
            {
                lblVypUbytek.Text = "-";
            }

            // 7. Oteplení
            double tAmbient = (comboBoxUlozeni != null && comboBoxUlozeni.SelectedIndex == 4) ? 20.0 : 30.0;
            double tprac = kabel.TpracstC > 0 ? kabel.TpracstC : 70.0;
            
            if (proudZat > 0)
            {
                // tn (provozní teplota pod jmenovitým zatížením)
                double tn = tAmbient + (tprac - tAmbient) * Math.Pow(proudZat / iz, 2.0);
                if (tn > tprac) tn = tprac; // limit na max pracovní teplotu
                
                // tm (mezní teplota při přetížení podle jističe In)
                double In = ZiskejNejblizsiJistic(proudZat);
                double I2 = 1.45 * In;
                double tm = tAmbient + (tprac - tAmbient) * Math.Pow(I2 / iz, 2.0);
                
                lblVypTeplota.Text = $"tn={tn:0.#} °C (provoz) / tm={tm:0.#} °C (přetížení {In}A)";
            }
            else
            {
                lblVypTeplota.Text = "-";
            }

            // Pomocná funkce pro formátování zkratových proudů
            string FormatCurrent(double amps) => amps >= 1000 ? $"{amps/1000.0:0.##} kA" : $"{amps:0} A";

            // 8. Zkrat sítě (3f / 1f) a tepelná odolnost kabelu
            if (delka > 0)
            {
                double ik3 = kabel.VypoctiZkrat3f(delka, napeti);
                double ik1 = kabel.VypoctiZkrat1f(delka, napeti);
                lblVypZkrat.Text = $"{FormatCurrent(ik3)} / {FormatCurrent(ik1)}";
            }
            else
            {
                lblVypZkrat.Text = "-";
            }

            bool jeHlinik = rbAl != null && rbAl.Checked;
            double zkrat01 = kabel.VypoctiZkratovyProud(jeHlinik, 0.1);
            double zkrat1 = kabel.VypoctiZkratovyProud(jeHlinik, 1.0);
            lblVypTepelnaOdolnost.Text = $"{FormatCurrent(zkrat01)} (0.1s) / {FormatCurrent(zkrat1)} (1s)";

            // 9. Impedance smyčky
            if (delka > 0)
            {
                double z = kabel.VypoctiImpedanciSmycky(delka);
                lblVypImpedance.Text = $"{z:0.###} Ω";
            }
            else
            {
                lblVypImpedance.Text = "-";
            }
        }

        private double ZiskejNejblizsiJistic(double proudA)
        {
            double[] standardniJistice = { 6, 10, 13, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 400, 630 };
            foreach (double j in standardniJistice)
            {
                if (j >= proudA) return j;
            }
            return proudA;
        }

        private void VynulujVypocty()
        {
            if (lblVypPrikon == null || lblVypNapeti == null || lblVypProud == null || lblVypDelka == null ||
                lblVypUbytek == null || lblVypZatizitelnost == null || lblVypTeplota == null || 
                lblVypZkrat == null || lblVypTepelnaOdolnost == null || lblVypImpedance == null) return;

            lblVypPrikon.Text = "-";
            lblVypNapeti.Text = "-";
            lblVypProud.Text = "-";
            lblVypDelka.Text = "-";
            lblVypUbytek.Text = "-";
            lblVypZatizitelnost.Text = "-";
            lblVypTeplota.Text = "-";
            lblVypZkrat.Text = "-";
            lblVypTepelnaOdolnost.Text = "-";
            lblVypImpedance.Text = "-";
        }
    }
}


