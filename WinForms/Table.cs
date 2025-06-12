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
            dataGridView1.AutoGenerateColumns = true;
            //dataGridView1.AutoGenerateColumns = false; // Vypnout automatické generování sloupců

            // Po připojení datového zdroje nahradíme sloupec Stav za ComboBox
            dataGridView1.DataSourceChanged += (s, e) =>
            {
                // Odstranit všechny sloupce, které nepotřebuješ
                var namesToRemove = new[] { "Druh", "DruhEnum", "Tag" };
                foreach (var name in namesToRemove)
                {
                    if (dataGridView1.Columns.Contains(name))
                        dataGridView1.Columns.Remove(name);
                }
            };

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
            string[] columnsToHide = 
            ["Pocet", "Nic", "PruzezMM2", "AWG", "Delka", "Delkaft", "Vyvod", "Druh", "Radek", "Vodice", "Kabel", 
            "Motor", "Patro", "Vykres", "IsExist", "Bod", "IsExistElektro", "Otoceni", "BodElektro", "HP", "Id" ];

            foreach (string columnName in columnsToHide)
            {
                if (dataGridView1.Columns.Contains(columnName))
                    dataGridView1.Columns[columnName]?.Visible = false;
            }
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
            foreach (DataGridViewColumn column in dataGridView1.Columns)
                column.Visible = true;
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "All")
            {
                //dataGridView1.DataSource = new SortableBindingList<ZarizeniView>(Pole);
                dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
                return;
            }

            string vybranePatro = comboBox1.SelectedItem.ToString();

            var filtrovanaData = Pole
                .Where(z => z.Patro == vybranePatro)
                .ToList();

            //dataGridView1.DataSource = new SortableBindingList<ZarizeniView>(filtrovanaData);
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(filtrovanaData);
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

        //Přidat
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.DataBoundItem is Zarizeni z)
            {
                z.Apid = ExcelLoad.Apid(); // Přidá nový prázdný záznam do seznamu
                Pole.Add(z);

            }
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole); // Obnoví datový zdroj pro zobrazení nového záznamu
        }

        //Delete
        private void Button7_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();
            var source = (BindingList<Zarizeni>)dataGridView1.DataSource;
            source.RemoveAt(dataGridView1.SelectedRows[0].Index);
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "All")
            {
                //dataGridView1.DataSource = new SortableBindingList<ZarizeniView>(Pole);
                dataGridView1.DataSource = new SortableBindingList<Zarizeni>(Pole);
                return;
            }
            string vybranePatro = comboBox2.SelectedItem.ToString();

            var filtrovanaData = Pole
                .Where(z => z.Etapa == vybranePatro)
                .ToList();

            //dataGridView1.DataSource = new SortableBindingList<ZarizeniView>(filtrovanaData);
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(filtrovanaData);
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
