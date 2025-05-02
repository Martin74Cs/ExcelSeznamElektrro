using Aplikace.Sdilene;
using Aplikace.Tridy;
using Aplikace.Upravy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace WinForms
{
    public partial class Table : Form
    {
        private List<Zarizeni> Pole { get; set; } // obecný typ, nebo použij generický s omezením
        //private SortableBindingList<Zarizeni> DataBind;
        //private BindingSource SourceBind = new BindingSource();
        public Table(List<Zarizeni> Pole)
        {
            this.Pole = Pole;
            InitializeComponent();
            SetListBox();
            //upravená třída BindingList na SortableBindingList
            var DataBind = new SortableBindingList<Zarizeni>(Pole);
            dataGridView1.CellFormatting += dataGridView1_CellFormatting;
            //SourceBind.DataSource = DataBind;
            dataGridView1.DataSource = DataBind;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv == null || dgv.Rows[e.RowIndex].DataBoundItem == null)
                return;

            Type type = typeof(Zarizeni);
            PropertyInfo[]vlastnosti = type.GetProperties();
            
            var Text = vlastnosti.Select(x => x.Name).ToArray();

            if (Text.Contains(dgv.Columns[e.ColumnIndex].Name) && !type.GetProperty(dgv.Columns[e.ColumnIndex].Name).CanWrite) // název sloupce ve zdroji dat
            {
                e.CellStyle.BackColor = Color.LightGray;
                dgv.Columns[dgv.Columns[e.ColumnIndex].Name].ReadOnly = true;
            }
        }

        public void SetListBox()
        {
            dataGridView1.AutoGenerateColumns = true;

            // Umožnit přidávání/smazání
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
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
            //Pole.AddProud();
            Pole.AddKabelCyky(1.5);
            dataGridView1.Refresh(); // obnoví zobrazení v datagridu
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
