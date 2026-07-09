using Aplikace.Sdilene;
using Aplikace.Tridy;
using Knihovna.Shared.Tridy;
using System;

namespace WinForms
{
    public partial class Shoda : Form
    {
        private List<Zarizeni> Strojni { get; set; } = []; // obecný typ, nebo použij generický s omezením
        private List<Zarizeni> Elektro { get; set; } = []; // obecný typ, nebo použij generický s omezením
        public Shoda(List<Zarizeni> strojni, List<Zarizeni> elektro) {
            this.Strojni = strojni;
            this.Elektro = elektro;
            InitializeComponent();
            SetListBox(dataGridView1);
            SetListBox(dataGridView2);
            //upravená třída BindingList na SortableBindingList

            var StrojniDataBind = new SortableBindingList<Zarizeni>();
            dataGridView1.DataSource = StrojniDataBind;
            // Skrýt některé sloupce

            var ElektroDataBind = new SortableBindingList<Zarizeni>(Elektro);
            dataGridView2.DataSource = ElektroDataBind;

            SkrytSloupce(dataGridView1);
            SkrytSloupce(dataGridView2);
        }

        private static void SkrytSloupce(DataGridView data) {
            string[] sloupceKeSkryti = [
                "Patro", "HP", "Delka", "IsExist", "IsExistElektro", "Bod", "BodElektro",
                "Nic", "AWG", "Delkaft", "PrurezMM2", "Rozvadec", "RozvadecCislo",
                "RozvadecOznačení", "Kabel", "Motor", "Vykres", "Vodice"
            ];
            foreach (var colName in sloupceKeSkryti)
            {
                var col = data.Columns[colName];
                if (col != null)
                {
                    col.Visible = false;
                }
            }
        }


        private void Shoda_Load(object sender, EventArgs e)
        {

        }


        private void DataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                //// Získání hodnoty buňky
                //// Zde můžete provést akci s hodnotou buňky, například ji zobrazit v MessageBoxu

                //{
                //    {

                //        // Zobrazíme druhý formulář jako modální dialog
                //    }

                    //funguje potom zapnout
                    //{
                    //    //if (index >= 0)
                    //    //{
                    //}
                    //else
                    //{ 
                    //}
                //}


            }
        }

        public static void SetListBox(DataGridView data)
        {
            data.AutoGenerateColumns = true;

            // Umožnit přidávání/smazání
            data.AllowUserToAddRows = true;
            data.AllowUserToDeleteRows = true;
        }

        private void DataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            var ShodaStrojni = new List<Zarizeni>();
            if (dataGridView2.CurrentRow?.DataBoundItem is Zarizeni selectedElektro) {
                if(selectedElektro.Tag.Length < 2) { return; }
                ShodaStrojni = [.. Strojni.Where(x => x.Tag.Contains(selectedElektro.Tag[..^1]))];
                if(ShodaStrojni.Count < 1 )
                    ShodaStrojni = [.. Strojni.Where(x => x.Tag.Contains(selectedElektro.Tag[..^2]))];
                        if (ShodaStrojni.Count < 1)
                            ShodaStrojni = [.. Strojni.Where(x => x.Tag.Contains(selectedElektro.Tag[..^3]))];
            }
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(ShodaStrojni);
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //přenos dat --- dole1 -> nahoru2
            if (dataGridView1.CurrentRow?.DataBoundItem is Zarizeni selectedStrojni &&
                dataGridView2.CurrentRow?.DataBoundItem is Zarizeni selectedElektro) {
                selectedElektro.Popis = selectedStrojni.Popis;
                selectedElektro.Radek = selectedStrojni.Radek;
                selectedElektro.Tag = selectedStrojni.Tag;
                selectedElektro.Menic = selectedStrojni.Menic;
                selectedElektro.Prikon = selectedStrojni.Prikon;
                selectedElektro.BalenaJednotka = selectedStrojni.BalenaJednotka;
                selectedElektro.Napeti = selectedStrojni.Napeti;

                var targetRow = dataGridView2.CurrentRow;
                if (targetRow != null)
                {
                    targetRow.DefaultCellStyle.BackColor = Color.LightGreen;
                    targetRow.DefaultCellStyle.ForeColor = Color.Black;
                }

                // Aby DataGridView vykreslil změny
            }
        }
    }
    
}
 



