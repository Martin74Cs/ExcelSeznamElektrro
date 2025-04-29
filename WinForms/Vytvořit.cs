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

namespace WinForms
{
    public partial class Vytvořit : Form
    {
        public Vytvořit()
        {
            InitializeComponent();
        }

        private void Vytvořit_Load(object sender, EventArgs e)
        {
            //Console.SetOut(new ListBoxWriter(listBox1));
        }

        //public class ListBoxWriter(ListBox listBox) : TextWriter
        //{
        //    private readonly ListBox _listBox = listBox;
        //    private readonly SynchronizationContext _context = SynchronizationContext.Current;

        //    public override Encoding Encoding => Encoding.UTF8;

        //    public override void WriteLine(string value)
        //    {
        //        //_context.Post(_ => _listBox.Items.Add(value), null);
        //        _context.Post(_ =>
        //        {
        //            _listBox.Items.Add(value);
        //            _listBox.TopIndex = _listBox.Items.Count - 1; // ← automatické scrollování dolů
        //        }, null);
        //    }
        //}

        private void Button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            //Close();
        }

        private async void Button7_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.VyvoritMotor());
        }

        private async void Button6_Click(object sender, EventArgs e)
        {
            //Menic
            //await Task.Run(() => LigthChem.VyvoritFM());
            var FM1 = Soubory.LoadFromCsv<Menic>(CestaKM);
            FM = new BindingList<Menic>(FM1);

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = KM;
            // Umožnit přidávání/smazání
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            //Motory open
            string cesta = Cesty.BasePath;
            cesta = System.IO.Path.Combine(cesta, "Data", "Motory");
            System.Diagnostics.Process.Start("explorer.exe", cesta);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Frekvenční měniče open
            string cesta = Cesty.BasePath;
            cesta = System.IO.Path.Combine(cesta, "Data", "FM.csv");
            System.Diagnostics.Process.Start("explorer.exe", cesta);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Stykače open
            System.Diagnostics.Process.Start("explorer.exe", CestaKM);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            //Stykače prevo
            //await Task.Run(() => LigthChem.VyvoritKM());

            var KM1 = Soubory.LoadFromCsv<Stykac>(CestaKM);
            KM = new BindingList<Stykac>(KM1);

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = KM;

            // Umožnit přidávání/smazání
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
        }

        private BindingList<Stykac> KM = [];
        private readonly string CestaKM = Path.Combine(Cesty.BasePath, "Data", "KM.csv");

        private BindingList<Menic> FM = [];
        private readonly string CestaFM = Path.Combine(Cesty.BasePath, "Data", "FM.csv");

        private void Button8_Click(object sender, EventArgs e)
        {
            Console.WriteLine($"Stykače uloženy jako Json a CSV");
            KM.ToList().SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            KM.ToList().SaveToCsv(CestaKM);
            dataGridView1.DataSource = null;
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            //save FM
            Console.WriteLine($"Menice uloženy jako Json a CSV");
            KM.ToList().SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            KM.ToList().SaveToCsv(CestaKM);
            dataGridView1.DataSource = null;
        }
    }
}
