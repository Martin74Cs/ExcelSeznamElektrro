using Aplikace.Sdilene;
using Aplikace.Tridy;
using Aplikace.Upravy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Aplikace.Tridy.Zarizeni;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using static WinForms.Table;

namespace WinForms
{
    public partial class Table: Form {
        private List<Zarizeni> Pole { get; set; }
        //private SortableBindingList<Zarizeni> DataBind;
        //private BindingSource SourceBind = new BindingSource();
        public Table(List<Zarizeni> Pole) {
            this.Pole = Pole;
            InitializeComponent();
            SetListBox();
            //upravená třída BindingList na SortableBindingList
            var DataBind = new SortableBindingList<Zarizeni>(Pole);
            dataGridView1.CellFormatting += dataGridView1_CellFormatting;
            dataGridView1.DataSource = DataBind;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            var dgv = sender as DataGridView;
            if(dgv == null || dgv.Rows[e.RowIndex].DataBoundItem == null)
                return;

            Type type = typeof(Zarizeni);
            PropertyInfo[] vlastnosti = type.GetProperties();

            var Text = vlastnosti.Select(x => x.Name).ToArray();

            if(Text.Contains(dgv.Columns[e.ColumnIndex].Name) && !type.GetProperty(dgv.Columns[e.ColumnIndex].Name).CanWrite) // název sloupce ve zdroji dat
            {
                e.CellStyle.BackColor = Color.LightGray;
                dgv.Columns[dgv.Columns[e.ColumnIndex].Name].ReadOnly = true;
            }
        }

        // Pomocná metoda pro získání popisu z enumu
        private string GetEnumDescription(Enum value) {
            var field = value.GetType().GetField(value.ToString());
            var attribute = (DescriptionAttribute)Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute));
            return attribute == null ? value.ToString() : attribute.Description;
        }

        public void SetListBox() {
            dataGridView1.AutoGenerateColumns = true;
            //dataGridView1.AutoGenerateColumns = false; // Vypnout automatické generování sloupců

            // Po připojení datového zdroje nahradíme sloupec Stav za ComboBox
            dataGridView1.DataSourceChanged += (s, e) => {
                var DruhColumn = dataGridView1.Columns["Druh"];
                DruhColumn.Visible = false;
                dataGridView1.Columns["DruhEnum"].Visible = false;
                //int index = stavColumn?.Index ?? 0;

                // Najdeme existující sloupec Stav
                //var stavColumn = dataGridView1.Columns["DruhEnum"];
                if(DruhColumn != null) {
                    // Získáme index sloupce
                    int columnIndex = DruhColumn.Index;

                    // Odstraníme původní sloupec
                    //dataGridView1.Columns.Remove(stavColumn);

                    // Vytvoříme seznam pro ComboBox s popisy
                    var Vyber = Enum.GetValues(typeof(Zarizeni.Druhy))
                    .Cast<Zarizeni.Druhy>().Select(s => new {
                        //Value = s.ToString(), // Ukládáme jako string
                        //Value = s, // Ukládáme jako string
                        Value = s.ToString(), // Ukládáme jako string
                        Display = GetEnumDescription(s) // Zobrazujeme popis
                    }).ToList();

                    // Vytvoříme nový ComboBox sloupec
                    DataGridViewComboBoxColumn comboBoxColumn = new DataGridViewComboBoxColumn {
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
            dataGridView1.AllowUserToAddRows = true;
            //dataGridView1.AllowUserToAddRows = false; // Zakázat přidávání prázdných řádků

            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter; // Umožnit editaci při kliknutí
        }

        private void Button1_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.Cancel;
        }

        private void Button2_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.OK;
        }

        private void Table_Load(object sender, EventArgs e) {

        }

        private void Button3_Click(object sender, EventArgs e) {
            //Proud
            if(Pole == null) return;
            Pole.AddProud();
            dataGridView1.Refresh(); // obnoví zobrazení v datagridu
        }

        private void Button4_Click(object sender, EventArgs e) {
            //průřez
            if(Pole == null) return;
            //Strojni.AddProud();
            Pole.AddKabelCyky(1.6);
            dataGridView1.Refresh(); // obnoví zobrazení v datagridu
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e) {
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

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e) {

        }

        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e) {
            var dgv = sender as DataGridView;
            if(dgv == null || dgv.CurrentCell == null || dgv.CurrentCell.RowIndex < 0)
                return;

            // Najdeme index sloupce "Stav"
            int stavColumnIndex = -1;
            foreach(DataGridViewColumn column in dgv.Columns) {
                if(column.Name == "Druh") {
                    stavColumnIndex = column.Index;
                    break;
                }
            }

            if(stavColumnIndex >= 0) {
                // Nastavíme aktuální buňku na sloupec "Stav" v aktuálním řádku
                //dgv.CurrentCell = dgv[stavColumnIndex, dgv.CurrentCell.RowIndex];
                dgv.BeginEdit(true);

                if(dgv.EditingControl is DataGridViewComboBoxEditingControl comboBox) {
                    comboBox.DroppedDown = true;
                }
            }
        }

        private void button5_Click(object sender, EventArgs e) {
            //Skrýtsloupce
            dataGridView1.Columns["PID"].Visible = false;
            dataGridView1.Columns["Pocet"].Visible = false;
            dataGridView1.Columns["Nic"].Visible = false;
            //dataGridView1.Columns["Proud"].Visible = false; 
            dataGridView1.Columns["PruzezMM2"].Visible = false;
            dataGridView1.Columns["AWG"].Visible = false;
            dataGridView1.Columns["Delka"].Visible = false;
            dataGridView1.Columns["Delkaft"].Visible = false;
            dataGridView1.Columns["Vyvod"].Visible = false;
            dataGridView1.Columns["Druh"].Visible = false;
            //dataGridView1.Columns["Napeti"].Visible = false; 
            dataGridView1.Columns["Radek"].Visible = false;
            dataGridView1.Columns["Vodice"].Visible = false;
            dataGridView1.Columns["Kabel"].Visible = false;
            dataGridView1.Columns["Motor"].Visible = false;
            dataGridView1.Columns["Patro"].Visible = false;
            dataGridView1.Columns["Vykres"].Visible = false;
            dataGridView1.Columns["IsExist"].Visible = false;
            dataGridView1.Columns["Bod"].Visible = false;
            dataGridView1.Columns["IsExistElektro"].Visible = false;
            dataGridView1.Columns["Otoceni"].Visible = false;
            dataGridView1.Columns["BodElektro"].Visible = false;
            dataGridView1.Columns["HP"].Visible = false;
            dataGridView1.Columns["Id"].Visible = false;
        }

        private void button6_Click(object sender, EventArgs e) {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
                column.Visible = true;
        }
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
