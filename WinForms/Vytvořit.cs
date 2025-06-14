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
    public partial class Vytvořit: Form {
        public Vytvořit() {
            InitializeComponent();
        }

        private void Vytvořit_Load(object sender, EventArgs e) {
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

        public void SetListBox() {
            dataGridView1.AutoGenerateColumns = true;

            // Umožnit přidávání/smazání
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
        }

        private void Button1_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            //Close();
        }

        private void Button7_Click(object sender, EventArgs e) {
            //Open motory
            //await Task.Run(() => LigthChem.VyvoritMotor());

            //var Motory1 = Soubory.LoadFromCsv<Motor>(CestaMotor);
            var Motory1 = Soubory.LoadJsonList<Motor>(CestaMotor);
            label2.Text = "Cesta = " + CestaMotor;
            Motor = new BindingList<Motor>(Motory1);
            dataGridView1.DataSource = Motor;
            SetListBox();
        }

        private void Button6_Click(object sender, EventArgs e) {
            //Open Menic
            //await Task.Run(() => LigthChem.VyvoritFM());
            var FM1 = Soubory.LoadFromCsv<Menic>(CestaFM);
            label2.Text = "Cesta = " + CestaFM;
            FM = new BindingList<Menic>(FM1);
            SetListBox();
            dataGridView1.DataSource = FM;
        }

        private void Button2_Click(object sender, EventArgs e) {
            //Motory open
            System.Diagnostics.Process.Start("explorer.exe", CestaMotor);
        }

        private void Button3_Click(object sender, EventArgs e) {
            //Frekvenční měniče open
            System.Diagnostics.Process.Start("explorer.exe", CestaFM);
        }

        private void Button4_Click(object sender, EventArgs e) {
            //Stykače open
            System.Diagnostics.Process.Start("explorer.exe", CestaKM);
        }

        private void Button5_Click(object sender, EventArgs e) {
            //Stykače opem
            //await Task.Run(() => LigthChem.VyvoritKM());
            var KM1 = Soubory.LoadFromCsv<Stykac>(CestaKM);
            label2.Text = "Cesta = " + CestaKM;
            KM = new BindingList<Stykac>(KM1);
            SetListBox();
            dataGridView1.DataSource = KM;
        }

        private string SaveCesta { get; set; }

        private BindingList<Stykac> KM = [];
        private readonly string CestaKM = Path.Combine(Cesty.Data, "KM.csv");

        private BindingList<Menic> FM = [];
        private readonly string CestaFM = Path.Combine(Cesty.Data, "FM.csv");

        private BindingList<Jistic> FA = [];
        private readonly string CestaJistic = Path.Combine(Cesty.Data, "Jištení", "Jističe3VA.csv");

        private BindingList<Motor> Motor = [];
        //private readonly string CestaMotor = Path.Combine(Cesty.Data, "Motory", "Motory.csv");

        private readonly string CestaMotor = Path.Combine(Cesty.Data, "Motory", "MotoryList.json");

        private void Button8_Click(object sender, EventArgs e) {
            //save Stykače
            Console.WriteLine($"Stykače uloženy jako Json a CSV");
            KM.ToList().SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            KM.ToList().SaveToCsv(CestaKM);
            dataGridView1.DataSource = null;
        }

        private void Button9_Click(object sender, EventArgs e) {
            //save FM
            Console.WriteLine($"Menice uloženy jako Json a CSV");
            KM.ToList().SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            KM.ToList().SaveToCsv(CestaKM);
            dataGridView1.DataSource = null;
        }

        private void Button10_Click(object sender, EventArgs e) {
            //save motory
            Console.WriteLine($"Menice uloženy jako Json a CSV");
            KM.ToList().SaveJsonList(Path.ChangeExtension(CestaMotor, ".json"));
            KM.ToList().SaveToCsv(CestaMotor);
            dataGridView1.DataSource = null;
        }

        private void Button11_Click(object sender, EventArgs e) {
            //Open oEz 3VA
            //var FA1 = Soubory.LoadFromCsv<Jistic>(CestaJistic);
            var FA1 = Soubory.LoadJsonList<Jistic>(Path.ChangeExtension(CestaJistic , ".json"));

            label2.Text = "Cesta = " + CestaJistic;
            FA = new BindingList<Jistic>(FA1);
            SetListBox();
            dataGridView1.DataSource = FA;

        }

        private void Vytvořit_FormClosing(object sender, FormClosingEventArgs e) {
            //save 
            Console.WriteLine($"Uloženy jako Json");
        }

        private void Button12_Click(object sender, EventArgs e) {
            //save  jističe
            Console.WriteLine($"Jističe uloženy jako Json.");
            FA.ToList().SaveJsonList(Path.ChangeExtension(CestaJistic, ".json"));
            dataGridView1.DataSource = null;
        }
    }
}
