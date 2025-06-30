using Aplikace.Sdilene;
using Aplikace.Tridy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinForms
{
    public partial class Shoda : Form
    {
        private List<Zarizeni> Strojni { get; set; } // obecný typ, nebo použij generický s omezením
        private List<Zarizeni> Elektro { get; set; } // obecný typ, nebo použij generický s omezením
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

        private void SkrytSloupce(DataGridView data) {
            data.Columns["Patro"].Visible = false;
            data.Columns["HP"].Visible = false;
            data.Columns["Delka"].Visible = false;
            data.Columns["IsExist"].Visible = false;
            data.Columns["IsExistElektro"].Visible = false;
            data.Columns["Bod"].Visible = false;
            data.Columns["BodElektro"].Visible = false;

            data.Columns["Nic"].Visible = false;
            data.Columns["AWG"].Visible = false;
            data.Columns["Delkaft"].Visible = false;

            data.Columns["PrurezMM2"].Visible = false;
            data.Columns["Rozvadec"].Visible = false;
            data.Columns["RozvadecCislo"].Visible = false;
            data.Columns["RozvadecOznačení"].Visible = false;
            data.Columns["Kabel"].Visible = false;
            data.Columns["Motor"].Visible = false;

            data.Columns["Vykres"].Visible = false;
            data.Columns["Vodice"].Visible = false;
            data.Columns["Motor"].Visible = false;
        }


        private void Shoda_Load(object sender, EventArgs e)
        {

        }


        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                //// Získání hodnoty buňky
                //var cellValue = data.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                //// Zde můžete provést akci s hodnotou buňky, například ji zobrazit v MessageBoxu
                //MessageBox.Show($"Hodnota buňky: {cellValue}");

                //foreach (var itemEl in Elektro.ToHashSet())
                //{
                //    var ShodaStrojni = Strojni.Where(x => x.Tag.Contains(itemEl.Tag[..^1])).ToList();
                //    if (ShodaStrojni.Count() > 1)
                //    {
                //        Console.WriteLine($"Shody Tagu - {ShodaStrojni.Count} kusy ");

                //        // Zobrazíme druhý formulář jako modální dialog
                //    }

                    //funguje potom zapnout
                    //var ShodaTag = Data.Where(x => x.Tag == item.Tag).ToList();
                    //if (ShodaTag.Count() == 1)
                    //{
                    //    var Jeden = ShodaTag.First();
                    //    Console.WriteLine($"Shoda je jedna - Doplněny pouze prázdné bunky ");
                    //    //var index = Data.IndexOf(Data.FirstOrDefault(x => x.Tag == item.Tag));
                    //    //if (index >= 0)
                    //    //{
                    //    item.Prikon = string.IsNullOrEmpty(item.Prikon) ? Jeden.Prikon : item.Prikon;
                    //    item.Menic = string.IsNullOrEmpty(item.Menic) ? Jeden.Menic : item.Menic;
                    //    item.BalenaJednotka = string.IsNullOrEmpty(item.BalenaJednotka) ? Jeden.BalenaJednotka : item.BalenaJednotka;
                    //    item.Pocet = item.Pocet == 0 ? Jeden.Pocet : item.Pocet;
                    //    item.Popis = string.IsNullOrEmpty(item.Popis) ? Jeden.Popis : item.Popis;
                    //    item.Radek = item.Radek == 0 ? Jeden.Radek : item.Radek;
                    //    item.Tag = string.IsNullOrEmpty(item.Tag) ? Jeden.Tag : item.Tag;
                    //    item.Napeti = string.IsNullOrEmpty(item.Napeti) ? Jeden.Napeti : item.Napeti;
                    //}
                    //else
                    //{ 
                    //    Console.WriteLine($"Kontrola - počet shod {ShodaTag.Count}");
                    //    item.Popis = $"KONTROLA - počet shod {ShodaTag.Count} ";
                    //}
                //}


            }
        }

        public void SetListBox(DataGridView data)
        {
            data.AutoGenerateColumns = true;

            // Umožnit přidávání/smazání
            data.AllowUserToAddRows = true;
            data.AllowUserToDeleteRows = true;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            var ShodaStrojni = new List<Zarizeni>();
            if (dataGridView2.CurrentRow?.DataBoundItem is Zarizeni selectedElektro) {
                if(selectedElektro.Tag.Length < 2) { return; }
                ShodaStrojni = Strojni.Where(x => x.Tag.Contains(selectedElektro.Tag[..^1])).ToList();
                if(ShodaStrojni.Count < 1 )
                    ShodaStrojni = Strojni.Where(x => x.Tag.Contains(selectedElektro.Tag[..^2])).ToList();
                        if (ShodaStrojni.Count < 1)
                            ShodaStrojni = Strojni.Where(x => x.Tag.Contains(selectedElektro.Tag[..^3])).ToList();
            }
            dataGridView1.DataSource = new SortableBindingList<Zarizeni>(ShodaStrojni);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
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

                // Najdeme instanci ve zdroji dat (předpoklad: BindingList nebo jiný upravitelný seznam)
                //if(dataGridView2.DataSource is BindingList<Zarizeni> elektro) {
                //    int index = dataGridView2.CurrentRow.Index;

                //    // Nahraďme celý objekt nebo jen jeho vlastnosti (záleží, jak s tím pracuješ dál)
                //    elektro[index] = new Zarizeni {
                //        Radek = selectedElektro.Radek,
                //        Prikon = selectedElektro.Prikon,
                //        Menic = selectedElektro.Menic,
                //        BalenaJednotka = selectedElektro.BalenaJednotka,
                //        Popis = selectedElektro.Popis,
                //        Tag = selectedElektro.Tag
                //        // přidej další vlastnosti podle potřeby
                //    };
                //}
                // Aby DataGridView vykreslil změny
                //dataGridView2.Refresh();
            }
        }
    }
    
}
 


