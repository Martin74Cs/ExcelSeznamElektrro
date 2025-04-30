using Aplikace.Sdilene;
using Aplikace.Tridy;
using Aplikace.Upravy;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace WinForms
{
    public partial class Table : Form
    {
        private List<Zarizeni> Pole { get; set; } // obecný typ, nebo použij generický s omezením

        public Table(List<Zarizeni> Pole)
        {
            this.Pole = Pole;
            InitializeComponent();
            SetListBox();
            var DataBind = new BindingList<Zarizeni>(Pole);
            dataGridView1.DataSource = DataBind;
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
