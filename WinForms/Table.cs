using Aplikace.Sdilene;
using Aplikace.Upravy;
using DocumentFormat.OpenXml.Drawing.Charts;
using Knihovna;
using Knihovna.Excel;
using Knihovna.Tridy;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Reflection;

namespace WinForms
{
    public partial class Table : Form
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public List<Zarizeni> PoleOut { get; set; }

        private List<Zarizeni> Pole { get; set; }

        /// <summary>Uchovává aktivní vlastní filtry pro zobrazení řádků.</summary>
        private List<FilterRule>? _customFilters = null;

        //private SortableBindingList<Popis> DataBind;
        //private BindingSource SourceBind = new BindingSource();
        public Table(List<Zarizeni> Pole)
        {
            this.Pole = Pole;
            InitializeComponent();
            var defaultColumns = new[] {
                nameof(Zarizeni.Tag),
                nameof(Zarizeni.Predmet),
                nameof(Zarizeni.Popis),
                nameof(Zarizeni.Druh),
                nameof(Zarizeni.Typ),
                nameof(Zarizeni.Menic),
                nameof(Zarizeni.Napeti),
                nameof(Zarizeni.Prikon),
                nameof(Zarizeni.Rozvadec),
                nameof(Zarizeni.RozvadecCislo),
                nameof(Zarizeni.Vyvod),
                //nameof(Zarizeni.Kabel),
                nameof(Zarizeni.SeznamKabelu)
            };
            SetListBox(defaultColumns);
            //upravená třída BindingList na SortableBindingList
            var DataBind = new SortableBindingList<Zarizeni>(Pole);
            dataGridView1.CellFormatting += DataGridView1_CellFormatting;
            dataGridView1.DataSource = DataBind;

            var zaklad = Pole.Select(z => z.Patro).Distinct().OrderBy(x => x).ToList();
            zaklad.Add("All");
            comboBox1.DataSource = zaklad; comboBox1.SelectedIndex = zaklad.Count - 1;

            var Etapa = Pole.Select(z => z.Etapa).Distinct().OrderBy(x => x).ToList();
            Etapa.Add("All");
            comboBox2.DataSource = Etapa; comboBox2.SelectedIndex = Etapa.Count - 1;

            var Rozvadec = Pole.Select(z => z.RozvadecOznačení).Distinct().OrderBy(x => x).ToList();
            Rozvadec.Add("All");
            comboBox3.DataSource = Rozvadec; comboBox3.SelectedIndex = Rozvadec.Count - 1;

            var PID = Pole.Select(z => z.Pid).Distinct().OrderBy(x => x).ToList();
            PID.Add("All");
            comboBox4Pid.DataSource = PID; comboBox4Pid.SelectedIndex = PID.Count - 1;

            var boolVyber = new List<string> { "All", "True", "False" };
            comboBoxIsExist.DataSource = new List<string>(boolVyber);
            comboBoxIsExist.SelectedIndex = 0;
            comboBoxIsExistElektro.DataSource = new List<string>(boolVyber);
            comboBoxIsExistElektro.SelectedIndex = 0;

            // Propojíme výběr v DataGridView se zobrazením v PropertyGridu
            dataGridView1.SelectionChanged += (s, e) =>
            {
                if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.DataBoundItem is Zarizeni z)
                {
                    propertyGrid1.SelectedObject = z;
                }
                else
                {
                    propertyGrid1.SelectedObject = null;
                }
            };

            // Když v PropertyGridu dojde ke změně hodnoty, překreslíme DataGridView
            propertyGrid1.PropertyValueChanged += (s, e) =>
            {
                dataGridView1.Refresh();
            };
        }

        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || sender is not DataGridView dgv || dgv.Rows[e.RowIndex].DataBoundItem == null)
                return;

            Type type = typeof(Zarizeni);
            PropertyInfo[] vlastnosti = type.GetProperties();

            var Text = vlastnosti.Select(x => x.Name).ToArray();

            if (Text.Contains(dgv.Columns[e.ColumnIndex].Name) && !type.GetProperty(dgv.Columns[e.ColumnIndex].Name).CanWrite) // název sloupce ve zdroji dat
            {
                e.CellStyle.BackColor = Color.LightGray;
                dgv.Columns[dgv.Columns[e.ColumnIndex].Name].ReadOnly = true;
            }
        }

        // Pomocná metoda pro získání popisu z enumu
        private static string GetEnumDescription(Enum value)
        {
            var field = value.GetType().GetField(value.ToString());
            var attribute = (DescriptionAttribute)Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute));
            return attribute == null ? value.ToString() : attribute.Description;
        }

        //public void SetListBoxOld() {
        //    dataGridView1.AutoGenerateColumns = true;
        //    //dataGridView1.AutoGenerateColumns = false; // Vypnout automatické generování sloupců

        //    // Po připojení datového zdroje nahradíme sloupec Stav za ComboBox
        //    dataGridView1.DataSourceChanged += (s, e) => {
        //        var DruhColumn = dataGridView1.Columns["Druh"];
        //        DruhColumn.Visible = false;
        //        dataGridView1.Columns["DruhEnum"]?.Visible = false;
        //        //int index = stavColumn?.Index ?? 0;

        //        // Najdeme existující sloupec Stav
        //        //var stavColumn = dataGridView1.Columns["DruhEnum"];
        //        if(DruhColumn != null) {
        //            // Získáme index sloupce
        //            int columnIndex = DruhColumn.Index;

        //            // Odstraníme původní sloupec
        //            //dataGridView1.Columns.Remove(stavColumn);

        //            // Vytvoříme seznam pro ComboBox s popisy
        //            var Vyber = Enum.GetValues<Popis.Druhy>()
        //            .Cast<Popis.Druhy>().Select(s => new {
        //                //Value = s.ToString(), // Ukládáme jako string
        //                //Value = s, // Ukládáme jako string
        //                Value = s.ToString(), // Ukládáme jako string
        //                Display = GetEnumDescription(s) // Zobrazujeme popis
        //            }).ToList();

        //            // Vytvoříme nový ComboBox sloupec
        //            var comboBoxColumn = new DataGridViewComboBoxColumn {
        //                HeaderText = "Vyber",
        //                Name = "Vyber",
        //                DataPropertyName = "Druh", // Propojení s vlastností Druh v Popis
        //                //DataSource = Enum.GetValues(typeof(Popis.Druhy)), // Naplní ComboBox hodnotami z enumu
        //                DataSource = Vyber,

        //                ValueMember = "Value", // String hodnota pro vlastnost Druh
        //                DisplayMember = "Display", // Zobrazení popisu
        //                ValueType = typeof(string)
        //                //ValueType = typeof(Popis.Druhy), // Zajistí správný typ hodnot
        //            };

        //            // Vložíme ComboBox sloupec na původní pozici
        //            dataGridView1.Columns.Insert(columnIndex, comboBoxColumn);
        //        }
        //    };

        //    // Přidání sloupce s ComboBoxem pro enum Stav   
        //    //DataGridViewComboBoxColumn comboBoxColumn = new DataGridViewComboBoxColumn
        //    //{
        //    //    HeaderText = "Druh",
        //    //    Name = "Druh",
        //    //    DataPropertyName = "Druh", // Propojení s vlastností Stav v Popis
        //    //    DataSource = Enum.GetValues(typeof(Popis.Druhy)), // Naplní ComboBox hodnotami z enumu
        //    //    ValueType = typeof(Popis.Druhy) // Zajistí správný typ hodnot
        //    //};
        //    //dataGridView1.Columns.Add(comboBoxColumn);

        //    // Umožnit přidávání/smazání
        //    //dataGridView1.AllowUserToAddRows = true;
        //    dataGridView1.AllowUserToAddRows = false; // Zakázat přidávání prázdných řádků

        //    dataGridView1.AllowUserToDeleteRows = true;
        //    dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter; // Umožnit editaci při kliknutí
        //}

        public void SetListBox(string[] propertyNames)
        {
            //dataGridView1.AutoGenerateColumns = true;
            dataGridView1.AutoGenerateColumns = false; // Vypnout automatické generování sloupců
            dataGridView1.Columns.Clear(); // důležité – vyčistí dříve vygenerované sloupce

            //"Druh"

            //var zarizeni = new Popis();

            // Přidáš sloupce ručně:
            foreach (var propertyName in propertyNames)
            {
                //if(name == "Druh") continue; // přeskočíme sloupec "Druh", ten bude přidán později

                var prop = typeof(Zarizeni).GetProperty(propertyName);
                if (prop == null)
                {
                    MessageBox.Show($"Vlastnost '{propertyName}' nebyla nalezena.");
                    continue;
                }

                Type propertyType = prop.PropertyType;

                // ENUM -> ComboBox
                if (propertyType.IsEnum)
                {
                    var comboColumn = new DataGridViewComboBoxColumn
                    {
                        Name = propertyName,
                        DataPropertyName = propertyName,
                        HeaderText = propertyName,
                        DataSource = Enum.GetValues(propertyType)
                    };

                    dataGridView1.Columns.Add(comboColumn);
                    continue;
                }

                if (propertyType == typeof(bool))
                {

                    //zaskrtávání pro bool
                    var checkColumn = new DataGridViewCheckBoxColumn
                    {
                        DataPropertyName = propertyName,
                        HeaderText = GetPropertyHeader(propertyName),
                        Name = propertyName
                    };
                    dataGridView1.Columns.Add(checkColumn);
                    continue;
                }

                //text ostatni
                var nameColumn = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = propertyName,
                    HeaderText = GetPropertyHeader(propertyName),
                    Name = propertyName
                };
                dataGridView1.Columns.Add(nameColumn);

            }

            //// Vytvoříme seznam pro ComboBox s popisy
            //var Vyber = Enum.GetValues<Popis.Druhy>()
            //.Cast<Popis.Druhy>().Select(s => new {
            //    Value = s.ToString(), // Ukládáme jako string
            //    Display = GetEnumDescription(s) // Zobrazujeme popis
            //}).ToList();

            //// Vytvoříme nový ComboBox sloupec
            //var comboBoxColumn = new DataGridViewComboBoxColumn {
            //    //HeaderText = "Vyber",
            //    //Name = "Vyber",
            //    //DataPropertyName = "Druh", 
            //    //DataSource = Vyber,
            //    HeaderText = GetPropertyHeader("Druh"),          // Nadpis sloupce
            //    Name = "Druh",                // Jméno sloupce
            //    DataPropertyName = "Druh",   // Vlastnost objektu Popis
            //    DataSource = Vyber,
            //    ValueMember = "Value",       // Skutečná hodnota (enum)
            //    DisplayMember = "Display",   // Co se zobrazí v roletce
            //    ValueType = typeof(Popis.Druhy)
            //};
            ////dataGridView1.Columns.Add(comboBoxColumn);
            //// Přidáme ComboBox sloupec na konec, nebo na určitou pozici
            //int position = propertyNames.Contains("Druh") ? propertyNames.IndexOf("Druh") : 1; // Najdeme index sloupce "Druh" v seznamu
            //dataGridView1.Columns.Insert(position, comboBoxColumn);

            // Umožnit přidávání/smazání
            //dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToAddRows = false; // Zakázat přidávání prázdných řádků

            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter; // Umožnit editaci při kliknutí
        }

        private static string GetPropertyHeader(string propertyName)
        {
            var prop = typeof(Zarizeni).GetProperty(propertyName);
            if (prop == null) return propertyName;

            var displayAttr = prop.GetCustomAttribute<System.ComponentModel.DataAnnotations.DisplayAttribute>();
            var jednotkyAttr = prop.GetCustomAttribute<JednotkyAttribute>();

            string header = displayAttr?.Name ?? propertyName;
            if (jednotkyAttr != null && !string.IsNullOrEmpty(jednotkyAttr.Text))
            {
                header += $" {jednotkyAttr.Text}";
            }
            return header;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void SpravaKabeluToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Zarizeni? vybraneZar = null;
            if (dataGridView1.CurrentRow != null)
            {
                vybraneZar = dataGridView1.CurrentRow.DataBoundItem as Zarizeni;
            }

            using var form = new FormKabely(Pole.OrderBy(x => x.Tag).ToList(), vybraneZar);
            try {
                form.ShowDialog();
            } catch(Exception) {
                Console.WriteLine("Divná chyba");
                throw;
            }
            dataGridView1.Refresh();
            propertyGrid1.Refresh();
        }

        private void HromadnaSpravaKabeluToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new FormVsechnyKabely(Pole);
            try
            {
                form.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Chyba při otevírání hromadné správy kabelů: {ex.Message}", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            dataGridView1.Refresh();
            propertyGrid1.Refresh();
        }

        private void PrirazeniKRozvadecumToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new FormRozvadece(Pole);
            form.ShowDialog(this);
            dataGridView1.Refresh();
            propertyGrid1.Refresh();
        }

        private void PrehledToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new Rozvaděč();
            form.ShowDialog(this);
        }

        private void Table_Load(object sender, EventArgs e)
        {

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            //Proud
            if (Pole == null) return;
            Pole.AddProud();
            dataGridView1.Refresh(); // obnoví zobrazení v datagridu
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            //průřez
            if (Pole == null) return;
            //Strojni.AddProud();
            //Pole.AddKabelCyky(1.6);
            //Pole.AddKabelCyky(2);
            dataGridView1.Refresh(); // obnoví zobrazení v datagridu
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //var dgv = sender as DataGridView;
            //if(dgv == null || e.RowIndex < 0 || e.ColumnIndex < 0)
            //    return;

            //// Zkontrolujeme, zda kliknutý sloupec je "Stav"
            //if(dgv.Columns[e.ColumnIndex].Name == "Stav") {
            //    // Aktivujeme editovací režim pro buňku
            //    dgv.CurrentCell = dgv[e.ColumnIndex, e.RowIndex];
            //    dgv.BeginEdit(true);

            //    // Otevřeme dropdown ComboBoxu
            //    if(dgv.EditingControl is DataGridViewComboBoxEditingControl comboBox) {
            //        comboBox.DroppedDown = true;
            //    }
            //}
        }

        private void DataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void DataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            if (sender is not DataGridView dgv || dgv.CurrentCell == null || dgv.CurrentCell.RowIndex < 0)
                return;

            // Najdeme index sloupce "Stav"
            int stavColumnIndex = -1;
            //foreach(DataGridViewColumn column in dgv.Columns) {
            //    if(column.Name == "Druh") {
            //        stavColumnIndex = column.Index;
            //        break;
            //    }
            //}

            if (stavColumnIndex >= 0)
            {
                // Nastavíme aktuální buňku na sloupec "Stav" v aktuálním řádku
                //dgv.CurrentCell = dgv[stavColumnIndex, dgv.CurrentCell.RowIndex];
                dgv.BeginEdit(true);

                if (dgv.EditingControl is DataGridViewComboBoxEditingControl comboBox)
                {
                    comboBox.DroppedDown = true;
                }
            }
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            _customFilters = null; // Vyčistíme pokročilé filtry
            var defaultColumns = new[] {
                nameof(Zarizeni.Tag),
                nameof(Zarizeni.Predmet),
                nameof(Zarizeni.Popis),
                nameof(Zarizeni.Druh),
                //nameof(Popis.Typ),
                nameof(Zarizeni.Menic),
                nameof(Zarizeni.Napeti),
                nameof(Zarizeni.Prikon),
                nameof(Zarizeni.Delka),
                nameof(Zarizeni.Rozvadec),
                nameof(Zarizeni.RozvadecCislo),
                nameof(Zarizeni.Vyvod),
                nameof(Zarizeni.Poznamka),
                //nameof(Popis.Kabel),
                //nameof(Popis.SeznamKabelu)
            };
            SetListBox(defaultColumns); // Obnoví sloupce v datagridu
            ObnovGrid(); // Načte kompletní seznam dat bez pokročilých filtrů
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            _customFilters = null; // Vyčistíme pokročilé filtry
            dataGridView1.Columns.Clear(); // důležité – vyčistí dříve vygenerované sloupce
            dataGridView1.AutoGenerateColumns = true;
            ObnovGrid(); // Znovu načte kompletní seznam dat
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            //if (box.Text == "All")
            //{
            //    dataGridView1.DataSource = new SortableBindingList<Popis>(Pole);
            //    return;
            //}
            //string vybranePatro = box.SelectedItem.ToString();
            //var filtrovanaData = Pole.Where(z => z.Patro == vybranePatro).ToList();

            ////dataGridView1.DataSource = new SortableBindingList<ZarizeniView>(filtrovanaData);
            //dataGridView1.DataSource = new SortableBindingList<Popis>(filtrovanaData);

            ObnovGrid();
        }

        private void Table_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dataGridView1.IsCurrentRowDirty)
            {
                dataGridView1.EndEdit();        // Ukončí editaci buňky
                dataGridView1.CurrentCell = null; // Vynutí commit řádku
                BindingContext[dataGridView1.DataSource].EndCurrentEdit(); // Vynutí uložení do seznamu
            }

        }

        private void DataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            // Ošetření, aby to neprobíhalo při načtení všech řádků znovu
            //if (e.RowIndex >= 0 && e.RowCount == 1)
            //{
            //    // Pokud přidání pochází od uživatele (ne automaticky), můžeme zachytit poslední řádek
            //    var posledni = dataGridView1.Rows[e.RowIndex];

            //    // Zde můžeš ověřit nebo vynutit uložení změn
            //    dataGridView1.EndEdit();

            //    // Můžeš například projít všechny řádky, nebo přistoupit k PoleDataBind a zkontrolovat, že nový řádek přibyl
            //    // nebo jen ohlásit změnu
            //    Console.WriteLine("Přidán nový řádek.");

            //    //pridat radek do pole
            //    Pole.Add(new Popis());

            //    dataGridView1.DataSource = new SortableBindingList<Popis>(Pole);
            //}
        }
        private Zarizeni _lastAddedOrEditedZarizeni = null;
        private string? _highlightedApid = null;
        //Přidat
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            // Zapamatujeme si aktuální pozici scrollbaru před přidáním, abychom zabránili skoku
            int scrollIndex = dataGridView1.FirstDisplayedScrollingRowIndex;

            if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.DataBoundItem is Zarizeni z)
            {
                var kopie = Zarizeni.Clone(z); // Zkopíruje aktuální řádek do nového záznamu
                kopie.Apid = ExcelLoad.Apid(); // Přidá se nový generovaný APID 
                kopie.SeznamKabelu = []; // Nulování kabelu aby se nekopírovaly
                //Všechno ostatní se kopíruje
                //kopie.Kabel = []; // Nulování kabelu aby se nekopírovaly

                // Vyhledáme skutečný index vybraného prvku v celkovém seznamu Pole
                int index = Pole.IndexOf(z);
                if (index >= 0)
                {
                    Pole.Insert(index + 1, kopie); // vložíme pod aktuální řádek
                }
                else
                {
                    Pole.Add(kopie);
                }
                _lastAddedOrEditedZarizeni = kopie;
                _highlightedApid = kopie.Apid;
            }
            else
            {
                //neni radek
                Pole.Add(new Zarizeni()); // vložíme na konec seznamu
            }

            ObnovGrid(zachovatPozici: false); // Vyvoláme obnovení bez automatického zachování

            // Výběr nově přidaného řádku a focus na první buňku
            if (!string.IsNullOrEmpty(_highlightedApid))
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.DataBoundItem is Zarizeni rowZarizeni && rowZarizeni.Apid == _highlightedApid)
                    {
                        row.Selected = true; // Vybereme řádek
                        row.DefaultCellStyle.BackColor = Color.LightSkyBlue;
                        dataGridView1.CurrentCell = row.Cells.Cast<DataGridViewCell>().FirstOrDefault(c => c.Visible); // Aktivní buňka
                        break;
                    }
                }
            }

            // Obnovíme scrollbar na původní pozici, nový řádek se objeví přirozeně hned pod ním
            if (scrollIndex >= 0 && scrollIndex < dataGridView1.Rows.Count)
            {
                dataGridView1.FirstDisplayedScrollingRowIndex = scrollIndex;
            }
        }

        private void ObnovGrid(bool zachovatPozici = true)
        {
            // Uložíme aktuálně vybraný prvek a pozici scrollbaru pro zachování plynulosti
            string? vybraneApid = null;
            int scrollIndex = -1;

            if (zachovatPozici && dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.DataBoundItem is Zarizeni z)
            {
                vybraneApid = z.Apid;
                scrollIndex = dataGridView1.FirstDisplayedScrollingRowIndex;
            }

            IEnumerable<Zarizeni> filtrovanaData = Pole;

            if (comboBox1.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.Patro == comboBox1.Text);

            if (comboBox2.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.Etapa == comboBox2.Text);

            if (comboBox3.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.RozvadecOznačení == comboBox3.Text);

            if (comboBox4Pid.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.Pid == comboBox4Pid.Text);

            // Fulltextové vyhledávání
            if (textBoxSearch != null && !string.IsNullOrWhiteSpace(textBoxSearch.Text))
            {
                string query = textBoxSearch.Text.Trim();
                filtrovanaData = filtrovanaData.Where(z =>
                    (z.Tag != null && z.Tag.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (z.Popis != null && z.Popis.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (z.Predmet != null && z.Predmet.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (z.Typ != null && z.Typ.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (z.Rozvadec != null && z.Rozvadec.Contains(query, StringComparison.OrdinalIgnoreCase))
                );
            }

            // Bool filtr IsExist
            if (comboBoxIsExist != null && comboBoxIsExist.Text != "All")
            {
                bool val = comboBoxIsExist.Text == "True";
                filtrovanaData = filtrovanaData.Where(z => z.IsExist == val);
            }

            // Bool filtr IsExistElektro
            if (comboBoxIsExistElektro != null && comboBoxIsExistElektro.Text != "All")
            {
                bool val = comboBoxIsExistElektro.Text == "True";
                filtrovanaData = filtrovanaData.Where(z => z.IsExistElektro == val);
            }

            // Aplikace vlastních filtrů (např. z FiltToolStripMenuItem_Click)
            if (_customFilters != null && _customFilters.Count > 0)
            {
                filtrovanaData = Soubory.ApplyFilter(filtrovanaData, _customFilters);
            }

            dataGridView1.DataSource = new SortableBindingList<Zarizeni>([.. filtrovanaData]);

            // Obnovíme výběr a pozici scrollu
            if (zachovatPozici && dataGridView1.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(vybraneApid))
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.DataBoundItem is Zarizeni rz && rz.Apid == vybraneApid)
                        {
                            dataGridView1.CurrentCell = row.Cells.Cast<DataGridViewCell>().FirstOrDefault(c => c.Visible);
                            break;
                        }
                    }
                }

                if (scrollIndex >= 0 && scrollIndex < dataGridView1.Rows.Count)
                {
                    dataGridView1.FirstDisplayedScrollingRowIndex = scrollIndex;
                }
            }
        }

        private void TextBoxSearch_TextChanged(object sender, EventArgs e)
        {
            ObnovGrid();
        }

        private void ComboBoxIsExist_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObnovGrid();
        }

        private void ComboBoxIsExistElektro_SelectedIndexChanged(object sender, EventArgs e)
        {
            ObnovGrid();
        }

        //Delete
        private void Button7_Click(object sender, EventArgs e)
        {
            var rowToSelect = dataGridView1.SelectedRows.Count > 0 ? dataGridView1.SelectedRows[0] : dataGridView1.CurrentRow;
            if (rowToSelect != null)
            {
                dataGridView1.EndEdit();

                if (rowToSelect.DataBoundItem is Zarizeni zarizeni)
                {
                    // Zapamatujeme si aktuální index a pozici scrollbaru před smazáním
                    int smazanyIndex = rowToSelect.Index;
                    int scrollIndex = dataGridView1.FirstDisplayedScrollingRowIndex;

                    Pole.Remove(zarizeni); // smažeme ze skutečného seznamu

                    //ObnovGrid(zachovatPozici: false); // Obnovíme grid bez automatického zachování
                    ObnovGrid(); // Obnovíme grid bez automatického zachování

                    // Po smazání vybereme řádek na stejné pozici, nebo předchozí řádek pokud šlo o poslední prvek
                    if (dataGridView1.Rows.Count > 0)
                    {
                        int novyIndex = Math.Min(smazanyIndex, dataGridView1.Rows.Count - 1);
                        if (novyIndex >= 0)
                        {
                            var row = dataGridView1.Rows[novyIndex];
                            dataGridView1.CurrentCell = row.Cells.Cast<DataGridViewCell>().FirstOrDefault(c => c.Visible);
                        }
                    }

                    //// Obnovíme pozici scrollbaru
                    //if (scrollIndex >= 0 && scrollIndex < dataGridView1.Rows.Count)
                    //{
                    //    dataGridView1.FirstDisplayedScrollingRowIndex = scrollIndex;
                    //}
                    //else if (dataGridView1.Rows.Count > 0)
                    //{
                    //    dataGridView1.FirstDisplayedScrollingRowIndex = Math.Max(0, dataGridView1.Rows.Count - 1);
                    //}
                }
            }
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            if (box.Text == "All")
            {
                dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
                return;
            }
            ObnovGrid(); // zachová aktuální filtry
            //string vybranePatro = box.SelectedItem.ToString();
            //var filtrovanaData = Pole.Where(z => z.Etapa == vybranePatro).ToList();
            //dataGridView1.DataSource = new SortableBindingList<Popis>(filtrovanaData);
        }

        private void ComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            if (box.Text == "All")
            {
                dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
                return;
            }
            ObnovGrid(); // zachová aktuální filtry
            //string vybranePatro = box.SelectedItem.ToString();
            //var filtrovanaData = Pole.Where(z => z.RozvadecOznačení == vybranePatro).ToList();
            //dataGridView1.DataSource = new SortableBindingList<Popis>(filtrovanaData);
        }

        private void ComboBox3_MouseClick(object sender, MouseEventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            box?.DroppedDown = true;
        }

        private void ComboBox2_MouseClick(object sender, MouseEventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            box?.DroppedDown = true;
        }

        private void ComboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            box?.DroppedDown = true;
        }

        private void ComboBox4Pid_SelectedIndexChanged(object sender, EventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            if (box.Text == "All")
            {
                dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
                return;
            }

            ObnovGrid(); // zachová aktuální filtry
            //string vybranePatro = box.SelectedItem.ToString();
            //var filtrovanaData = Pole.Where(z => z.PID == vybranePatro).ToList();
            //dataGridView1.DataSource = new SortableBindingList<Popis>(filtrovanaData);
        }

        private void ComboBox4Pid_MouseClick(object sender, MouseEventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            box?.DroppedDown = true;
        }

        // --- Událost pro obarvení řádku ---
        //private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        //{
        //    // Zkontrolujeme, zda máme nějaké APID k zvýraznění a zda je index řádku platný
        //    if (!string.IsNullOrEmpty(_highlightedApid) && e.RowIndex >= 0)
        //    {
        //        //// Získáme objekt, ke kterému je aktuální řádek vázán
        //        if (dataGridView1.Rows[e.RowIndex].DataBoundItem is Popis rowZarizeni)
        //        {
        //        //    // Porovnáme APID řádku s APID, které chceme zvýraznit
        //            if (rowZarizeni.Apid == _highlightedApid)
        //            {
        //                e.CellStyle.BackColor = Color.LightGreen; // Barva pro zvýrazněný řádek
        //                e.FormattingApplied = true; // Řekne DataGridView, že jsme barvu aplikovali
        //            }
        //        //    else
        //        //    {
        //        //        // Pokud řádek NENÍ ten, který má být obarven, resetujeme jeho barvu na výchozí
        //        //        e.CellStyle.BackColor = Color.Empty; // Reset na výchozí barvu (transparentní)
        //        //        e.FormattingApplied = true;
        //        //    }
        //        }
        //    }
        //    else
        //    {
        //        // Pokud _highlightedApid je null (nebo neplatný RowIndex),
        //        // ujistěte se, že žádný řádek není obarven a je nastaven na výchozí barvu.
        //        //if (e.RowIndex >= 0)
        //        //{
        //        //    e.CellStyle.BackColor = Color.Empty;
        //        //    e.FormattingApplied = true;
        //        //}
        //    }
        //}

        // Metoda pro explicitní odstranění zvýraznění (např. po uložení)
        public void ResetHighlight()
        {
            _highlightedApid = null; // Vymaže APID, které se má zvýraznit
            dataGridView1.Invalidate(); // Vynutí překreslení DataGridView (resetuje barvy)
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            _customFilters = null; // Vyčistíme pokročilé filtry
            var dataColumns = new[] {
                nameof(Zarizeni.Tag),
                nameof(Zarizeni.Predmet),
                nameof(Zarizeni.Popis),
                nameof(Zarizeni.Druh),
                nameof(Zarizeni.Typ),
                nameof(Zarizeni.PrikonStroj),
                nameof(Zarizeni.Prikon),
                nameof(Zarizeni.Napeti),
                nameof(Zarizeni.Menic),
                nameof(Zarizeni.Proud),
                nameof(Zarizeni.RozvadecOznačení),
                nameof(Zarizeni.PrurezMM2),
                //nameof(Zarizeni.Kabel),
                nameof(Zarizeni.SeznamKabelu)
            };
            SetListBox(dataColumns); // Obnoví sloupce v datagridu
            ObnovGrid(); // Načte kompletní seznam dat bez pokročilých filtrů
        }

        private void FiltToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Sloupce k zobrazení podle exportu ve Form1
            var namesToRemove = new[] {
                nameof(Zarizeni.Tag),
                nameof(Zarizeni.Predmet),
                nameof(Zarizeni.Popis),
                nameof(Zarizeni.Prikon),
                nameof(Zarizeni.Napeti),
                nameof(Zarizeni.Proud),
                nameof(Zarizeni.Menic),
                nameof(Zarizeni.BalenaJednotka),
                nameof(Zarizeni.Pozice)
            };
            SetListBox(namesToRemove);

            // Vytvoření filtrů jako u Form1
            _customFilters = [
            
                new(nameof(Zarizeni.Prikon), op: FilterOperator.IsNotNullOrEmpty),
                // Všechny položky, jejichž Příkon nezačíná na "—" ani "-"
                new(nameof(Zarizeni.Prikon), "—", op: FilterOperator.StartsWith, negate: true),
                new(nameof(Zarizeni.Prikon), "-", op: FilterOperator.StartsWith, negate: true),
            ];

            ObnovGrid(); // Aplikuje filtry na data
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Cesta = Soubory.ShowOpenFileDialog("Json soubor (*.json)|*.json", Informace.Instance.BasePath);
            if (File.Exists(Cesta))
            {
                Informace.Instance.SouborElektroJson = Cesta;
                Informace.Instance.Ulozit();
            }
            else
                Console.WriteLine("Soubor Nexistuje");
        }

        private void BezKWToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (var item in Pole)
            {
                //pokud neni čislo tak to smazat
                if (!double.TryParse(item.Prikon, out double _))
                {
                    Pole.Remove(item); // smažeme ze skutečného seznamu
                }
            }
            ObnovGrid(zachovatPozici: false); // Obnovíme grid bez automatického zachování
        }
    }

    public class SortableBindingList<T> : BindingList<T>
    {
        public SortableBindingList() : base() { }

        public SortableBindingList(IList<T> list) : base(list) { }

        private bool isSorted;
        private ListSortDirection sortDirection;
        private PropertyDescriptor sortProperty;

        protected override bool SupportsSortingCore => true;
        protected override bool IsSortedCore => isSorted;

        protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
        {
            var items = (List<T>)Items;
            items.Sort((x, y) =>
            {
                var xValue = prop.GetValue(x);
                var yValue = prop.GetValue(y);
                return direction == ListSortDirection.Ascending
                    ? Comparer.DefaultInvariant.Compare(xValue, yValue)
                    : Comparer.DefaultInvariant.Compare(yValue, xValue);
            });

            sortDirection = direction;
            sortProperty = prop;
            isSorted = true;
            OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
        }

        protected override void RemoveSortCore()
        {
            isSorted = false;
        }

        protected override PropertyDescriptor SortPropertyCore => sortProperty;
        protected override ListSortDirection SortDirectionCore => sortDirection;
    }
}
