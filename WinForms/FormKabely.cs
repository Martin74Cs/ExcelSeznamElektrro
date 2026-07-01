using Knihovna;
using Knihovna.Shared.Tridy;
using Knihovna.Tridy;
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

        public FormKabely(List<Zarizeni> seznamZarizeni, Zarizeni? predvoleneZarizeni = null) {
            InitializeComponent();
            _seznamZarizeni = seznamZarizeni;
            _predvoleneZarizeni = predvoleneZarizeni;
        }

        private void FormKabely_Load(object sender, EventArgs e) {
            _nacitani = true;

            // Naplnění seznamu značek
            comboBoxZnacka.DataSource = Enum.GetValues<KabelZnačka>();
            comboBoxZnacka.SelectedItem = KabelZnačka.WL;

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
        }

        private void ComboBoxZarizeni_SelectedIndexChanged(object sender, EventArgs e) {
            // Při změně zařízení zrušíme případný režim úpravy
            ResetFormNaNovaKabel();
            ObnovSeznamKabelu();
            NactiVychoziHodnotyZařízení();
        }

        private void ComboBoxZnacka_SelectedIndexChanged(object sender, EventArgs e) {
            // Automatický návrh označení jen pokud NEJSME v režimu úpravy
            if(_upravovanyKabel == null) {
                AutoSuggestOznaceni();
            }
        }

        private void ComboBoxFilter_SelectedIndexChanged(object sender, EventArgs e) {
            if(_nacitani) return;
            AplikujFiltryZarizeni();
        }

        private void AplikujFiltryZarizeni() {
            var predesleVybrane = comboBoxZarizeni.SelectedItem as Zarizeni;

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
                string vybranaEtapa = comboBoxFilterEtapa.SelectedItem.ToString();
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
            if(predesleVybrane != null && filtrovanySeznam.Contains(predesleVybrane)) {
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
                txtPrurez.Text = activeZar.PrurezMM2;
                txtDelka.Text = activeZar.Delka.ToString("0.##");
                txtPopis.Text = string.Empty;
            }
        }

        private static string GenerujUnikantiOznaceni(string znacka, Zarizeni activeZar) {
            int maxNum = 0;
            //foreach(var z in _seznamZarizeni) {
            foreach (var k in activeZar.SeznamKabelu)  {
             //   foreach (var k in _predvoleneZarizeni.SeznamKabelu) {
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
            //    string.Equals(k.Oznaceni, noveOznaceni, StringComparison.OrdinalIgnoreCase)));

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

                //UlozData();
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

                activeZar.SeznamKabelu.Add(trasa);

                //UlozData();
                ObnovSeznamKabelu();
            }
        }

        private void BtnSmazat_Click(object sender, EventArgs e) {
            if(comboBoxZarizeni.SelectedItem is not Zarizeni activeZar || dataGridViewKabely.CurrentRow == null)
                return;

            if(dataGridViewKabely.CurrentRow.DataBoundItem is Trasa vybranyKabel) {
                if(MessageBox.Show($"Opravdu chcete smazat kabel '{vybranyKabel.Oznaceni}'?", "Potvrzení", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    activeZar.SeznamKabelu.Remove(vybranyKabel);
                    //UlozData();
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
                txtPrurez.Text = activeZar.PrurezMM2;
                txtDelka.Text = activeZar.Delka.ToString("0.##");
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
                Delka = activeZar.Delka.ToString("0.##"),
                Patro = activeZar.Patro,
                Predmet = activeZar.Predmet
            };

            activeZar.SeznamKabelu.Add(trasa);

            //UlozData();
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
            //PridejRychlyKabel(prefix, "CYKY-O", "7", "2.5", "Ovládací skříňka");
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

        //private void UlozData() {
        //    try {
        //        //string Cesta = Path.Combine(Informace.Instance.BasePath, "Elektro.Data.Json");
        //        //_seznamZarizeni.SaveJsonList(Cesta);
        //    } catch(Exception ex) {
        //        MessageBox.Show($"Nepodařilo se uložit data do souboru: {ex.Message}", "Chyba ukládání", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

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
    }
}

