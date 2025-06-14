using Aplikace.Sdilene;
using Aplikace.Tridy;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinForms {
    public partial class Rozvaděč: Form {
        private List<Zarizeni> Data = [];
        public Rozvaděč() {
            InitializeComponent();
            //Data = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            var Vývody = Path.Combine(Cesty.Elektro, "Vývody.json");
            Data = Soubory.LoadJsonList<Zarizeni>(Vývody);
            Set(listViewCategories);
            Set(listViewProducts);
        }

        private void Set(ListView list) {
            list.View = View.Details;
            list.FullRowSelect = true;
            list.GridLines = true;
            list.Columns.Add("Rozvadeč", 60);
            list.Columns.Add("Popis", 200);
            list.Columns.Add("Tag", 60);
        }

        private void Rozvaděč_Load(object sender, EventArgs e) {

            foreach(var dat in Data.DistinctBy(x => x.RozvadecOznačení)) {
                var item = new ListViewItem(dat.RozvadecOznačení) {
                    //item.SubItems.Add(dat.RozvadecOznačení);
                    //item.SubItems.Add(dat.Popis);
                    //item.SubItems.Add(dat.Tag);
                    Tag = dat // Ulož celou instanci pro pozdější použití
                };
                listViewCategories.Items.Add(item);
            }
        }

        private void ListViewCategories_SelectedIndexChanged(object sender, EventArgs e) {
            if (listViewCategories.SelectedItems.Count > 0)
            {
                var item = listViewCategories.SelectedItems[0];
                if (item.Tag is Zarizeni zzz)
                {
                    var vvv = Data.Where(x => x.RozvadecOznačení == zzz.RozvadecOznačení).ToList();
                    //MessageBox.Show($"Test: {zzz.Tag}");
                    
                    listViewProducts.Items.Clear();
                    foreach(var dat in vvv) {
                        var lll = new ListViewItem(dat.RozvadecOznačení);
                        //lll.SubItems.Add(dat.RozvadecOznačení);
                        lll.SubItems.Add(dat.Popis);
                        lll.SubItems.Add(dat.Tag);
                        lll.Tag = dat; // Ulož celou instanci pro pozdější použití
                        listViewProducts.Items.Add(lll);
                    }
                }
            }
        }
    }
}
