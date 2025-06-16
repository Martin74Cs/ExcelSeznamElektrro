using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Tridy;
using Aplikace.Upravy;
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

        //private SortableBindingList<Zarizeni> DataBind;
        //private BindingSource SourceBind = new BindingSource();
        public Table(List<Zarizeni> Pole)
        {
            this.Pole = Pole;
            InitializeComponent();
            SetListBox();
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

            var PID = Pole.Select(z => z.PID).Distinct().OrderBy(x => x).ToList();
            PID.Add("All");
            comboBox4Pid.DataSource = PID; comboBox4Pid.SelectedIndex = PID.Count - 1;
        }

        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (sender is not DataGridView dgv || dgv.Rows[e.RowIndex].DataBoundItem == null)
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

        public void SetListBoxOld()
        {
            dataGridView1.AutoGenerateColumns = true;
            //dataGridView1.AutoGenerateColumns = false; // Vypnout automatické generování sloupců

            // Po připojení datového zdroje nahradíme sloupec Stav za ComboBox
            dataGridView1.DataSourceChanged += (s, e) =>
            {
                var DruhColumn = dataGridView1.Columns["Druh"];
                DruhColumn.Visible = false;
                dataGridView1.Columns["DruhEnum"]?.Visible = false;
                //int index = stavColumn?.Index ?? 0;

                // Najdeme existující sloupec Stav
                //var stavColumn = dataGridView1.Columns["DruhEnum"];
                if (DruhColumn != null)
                {
                    // Získáme index sloupce
                    int columnIndex = DruhColumn.Index;

                    // Odstraníme původní sloupec
                    //dataGridView1.Columns.Remove(stavColumn);

                    // Vytvoříme seznam pro ComboBox s popisy
                    var Vyber = Enum.GetValues<Zarizeni.Druhy>()
                    .Cast<Zarizeni.Druhy>().Select(s => new
                    {
                        //Value = s.ToString(), // Ukládáme jako string
                        //Value = s, // Ukládáme jako string
                        Value = s.ToString(), // Ukládáme jako string
                        Display = GetEnumDescription(s) // Zobrazujeme popis
                    }).ToList();

                    // Vytvoříme nový ComboBox sloupec
                    var comboBoxColumn = new DataGridViewComboBoxColumn
                    {
                        HeaderText = "Vyber",
                        Name = "Vyber",
                        DataPropertyName = "Druh", // Propojení s vlastností Druh v Zarizeni
                        //DataSource = Enum.GetValues(typeof(Zarizeni.Druhy)), // Naplní ComboBox hodnotami z enumu
                        DataSource = Vyber,

                        ValueMember = "Value", // String hodnota pro vlastnost Druh
                        DisplayMember = "Display", // Zobrazení popisu
                        ValueType = typeof(string)
                        //ValueType = typeof(Zarizeni.Druhy), // Zajistí správný typ hodnot
                    };

                    // Vložíme ComboBox sloupec na původní pozici
                    dataGridView1.Columns.Insert(columnIndex, comboBoxColumn);
                }
            };

            // Přidání sloupce s ComboBoxem pro enum Stav   
            //DataGridViewComboBoxColumn comboBoxColumn = new DataGridViewComboBoxColumn
            //{
            //    HeaderText = "Druh",
            //    Name = "Druh",
            //    DataPropertyName = "Druh", // Propojení s vlastností Stav v Zarizeni
            //    DataSource = Enum.GetValues(typeof(Zarizeni.Druhy)), // Naplní ComboBox hodnotami z enumu
            //    ValueType = typeof(Zarizeni.Druhy) // Zajistí správný typ hodnot
            //};
            //dataGridView1.Columns.Add(comboBoxColumn);

            // Umožnit přidávání/smazání
            //dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToAddRows = false; // Zakázat přidávání prázdných řádků

            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter; // Umožnit editaci při kliknutí
        }

        public void SetListBox()
        {
            var namesToRemove = new[] { "TagStroj", "Tag", "Predmet", "Popis", "Typ", "Pid", "Menic", "Prikon", "PrikonStroj", "Rozvadec", "RozvadecCislo", "RozvadecOznačení", "Nic", "Delka", "Vyvod", "Patro", "Vykres" };
            SetListBox(namesToRemove);
        }
        public void SetListBoxData()
        {
            var namesToRemove = new[] { "Tag", "Predmet", "Popis", "Typ", "Prikon", "Napeti", "Menic", "Proud", "RozvadecOznačení", "PruzezMM2" };
            SetListBox(namesToRemove);
        }

        public void SetListBox(string[] namesToRemove)
        {
            //dataGridView1.AutoGenerateColumns = true;
            dataGridView1.AutoGenerateColumns = false; // Vypnout automatické generování sloupců
            dataGridView1.Columns.Clear(); // důležité – vyčistí dříve vygenerované sloupce

            //"Druh"

            // Přidáš sloupce ručně:
            foreach (var name in namesToRemove)
            {
                var nameColumn = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = name,
                    HeaderText = name,
                    Name = name
                };
                dataGridView1.Columns.Add(nameColumn);

            }

            // Vytvoříme seznam pro ComboBox s popisy
            var Vyber = Enum.GetValues<Zarizeni.Druhy>()
            .Cast<Zarizeni.Druhy>().Select(s => new
            {
                Value = s.ToString(), // Ukládáme jako string
                Display = GetEnumDescription(s) // Zobrazujeme popis
            }).ToList();

            // Vytvoříme nový ComboBox sloupec
            var comboBoxColumn = new DataGridViewComboBoxColumn
            {
                //HeaderText = "Vyber",
                //Name = "Vyber",
                //DataPropertyName = "Druh", 
                //DataSource = Vyber,
                HeaderText = "Druh",          // Nadpis sloupce
                Name = "Druh",                // Jméno sloupce
                DataPropertyName = "Druh",   // Vlastnost objektu Zarizeni
                DataSource = Vyber,
                ValueMember = "Value",       // Skutečná hodnota (enum)
                DisplayMember = "Display",   // Co se zobrazí v roletce
                ValueType = typeof(Zarizeni.Druhy)
            };
            //dataGridView1.Columns.Add(comboBoxColumn);
            dataGridView1.Columns.Insert(4, comboBoxColumn);
            // Umožnit přidávání/smazání
            //dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToAddRows = false; // Zakázat přidávání prázdných řádků

            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter; // Umožnit editaci při kliknutí
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
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
            Pole.AddKabelCyky(1.6);
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
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (column.Name == "Druh")
                {
                    stavColumnIndex = column.Index;
                    break;
                }
            }

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
            SetListBox(); // Obnoví sloupce v datagridu

            //string[] columnsToHide =
            //["Pocet", "Nic", "PruzezMM2", "AWG", "Delka", "Delkaft", "Vyvod", "Druh", "Radek", "Vodice", "Kabel",
            //"Motor", "Patro", "Vykres", "IsExist", "Bod", "IsExistElektro", "Otoceni", "BodElektro", "HP", "Id" ];

            //foreach (string columnName in columnsToHide)
            //{
            //    if (dataGridView1.Columns.Contains(columnName))
            //        dataGridView1.Columns[columnName]?.Visible = false;
            //}
            ////Skrýtsloupce
            ////dataGridView1.Columns["PID"].Visible = false;
            //dataGridView1.Columns["Pocet"]?.Visible = false;
            //dataGridView1.Columns["Nic"]?.Visible = false;
            ////dataGridView1.Columns["Proud"].Visible = false; 
            //dataGridView1.Columns["PruzezMM2"]?.Visible = false;
            //dataGridView1.Columns["AWG"]?.Visible = false;
            //dataGridView1.Columns["Delka"]?.Visible = false;
            //dataGridView1.Columns["Delkaft"]?.Visible = false;
            //dataGridView1.Columns["Vyvod"]?.Visible = false;
            //dataGridView1.Columns["Druh"]?.Visible = false;
            ////dataGridView1.Columns["Napeti"].Visible = false; 
            //dataGridView1.Columns["Radek"]?.Visible = false;
            //dataGridView1.Columns["Vodice"]?.Visible = false;
            //dataGridView1.Columns["Kabel"]?.Visible = false;
            //dataGridView1.Columns["Motor"]?.Visible = false;
            //dataGridView1.Columns["Patro"]?.Visible = false;
            //dataGridView1.Columns["Vykres"]?.Visible = false;
            //dataGridView1.Columns["IsExist"]?.Visible = false;
            //dataGridView1.Columns["Bod"]?.Visible = false;
            //dataGridView1.Columns["IsExistElektro"]?.Visible = false;
            //dataGridView1.Columns["Otoceni"]?.Visible = false;
            //dataGridView1.Columns["BodElektro"]?.Visible = false;
            //dataGridView1.Columns["HP"]?.Visible = false;
            //dataGridView1.Columns["Id"]?.Visible = false;
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear(); // důležité – vyčistí dříve vygenerované sloupce
            dataGridView1.AutoGenerateColumns = true;
            //foreach (DataGridViewColumn column in dataGridView1.Columns)
            //    column.Visible = true;
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            //if (box.Text == "All")
            //{
            //    dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
            //    return;
            //}
            //string vybranePatro = box.SelectedItem.ToString();
            //var filtrovanaData = Pole.Where(z => z.Patro == vybranePatro).ToList();

            ////dataGridView1.DataSource = new SortableBindingList<ZarizeniView>(filtrovanaData);
            //dataGridView1.DataSource = new SortableBindingList<Zarizeni>(filtrovanaData);

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
            //    Pole.Add(new Zarizeni());

            //    dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
            //}
        }
        private Zarizeni _lastAddedOrEditedZarizeni = null;
        private string? _highlightedApid = null;
        //Přidat
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.DataBoundItem is Zarizeni z)
            {
                var kopie = Zarizeni.Clone(z); // Zkopíruje aktuální řádek do nového záznamu
                kopie.Apid = ExcelLoad.Apid(); // Přidá nový prázdný záznam do seznamu

                int index = dataGridView1.CurrentRow.Index;
                Pole.Insert(index + 1, kopie); // vložíme pod aktuální řádek
                _lastAddedOrEditedZarizeni = kopie;
                _highlightedApid = kopie.Apid;
            }
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole); // Obnoví datový zdroj pro zobrazení nového záznamu

            ObnovGrid(); // zachová aktuální filtry

            // Volitelné: Scroll na nově přidaný řádek a jeho výběr
            if (!string.IsNullOrEmpty(_highlightedApid))
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.DataBoundItem is Zarizeni rowZarizeni && rowZarizeni.Apid == _highlightedApid)
                    {
                        dataGridView1.FirstDisplayedScrollingRowIndex = row.Index;
                        //e.CellStyle.BackColor = Color.LightGreen;
                        row.Selected = true; // Volitelně: vyberte řádek
                        row.DefaultCellStyle.BackColor = Color.LightSkyBlue;
                        dataGridView1.CurrentCell = row.Cells[0]; // Fokus na buňku
                        break;

                    }
                }
            }
        }

        private void ObnovGrid()
        {
            IEnumerable<Zarizeni> filtrovanaData = Pole;

            if (comboBox1.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.Patro == comboBox1.Text);

            if (comboBox2.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.Etapa == comboBox2.Text);

            if (comboBox3.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.RozvadecOznačení == comboBox3.Text);

            if (comboBox4Pid.Text != "All")
                filtrovanaData = filtrovanaData.Where(z => z.PID == comboBox4Pid.Text);

            dataGridView1.DataSource = new SortableBindingList<Zarizeni>([.. filtrovanaData]);
        }

        //Delete
        private void Button7_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                dataGridView1.EndEdit();

                var zarizeni = dataGridView1.SelectedRows[0].DataBoundItem as Zarizeni;
                if (zarizeni != null)
                {
                    Pole.Remove(zarizeni); // smažeme ze skutečného seznamu
                    ObnovGrid(); // obnovíme zobrazení podle aktivních filtrů
                }
            }
            //dataGridView1.EndEdit();
            //var source = (BindingList<Zarizeni>)dataGridView1.DataSource;
            //source.RemoveAt(dataGridView1.SelectedRows[0].Index);
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            if (box.Text == "All")
            {
                dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
                return;
            }
            string vybranePatro = box.SelectedItem.ToString();
            var filtrovanaData = Pole.Where(z => z.Etapa == vybranePatro).ToList();
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(filtrovanaData);
        }

        private void ComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            var box = sender as ComboBox; // Získání ComboBoxu, který vyvolal událost
            if (box.Text == "All")
            {
                dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
                return;
            }
            string vybranePatro = box.SelectedItem.ToString();
            var filtrovanaData = Pole.Where(z => z.RozvadecOznačení == vybranePatro).ToList();
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(filtrovanaData);
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
            string vybranePatro = box.SelectedItem.ToString();
            var filtrovanaData = Pole.Where(z => z.PID == vybranePatro).ToList();
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(filtrovanaData);
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
        //        if (dataGridView1.Rows[e.RowIndex].DataBoundItem is Zarizeni rowZarizeni)
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

        private void button8_Click(object sender, EventArgs e)
        {
            SetListBoxData();
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
