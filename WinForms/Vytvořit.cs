
using Aplikace.Sdilene;
using Knihovna.Tridy;
using Knihovna.Sdilene;
using System.ComponentModel;
using Knihovna;

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
            DialogResult = DialogResult.Cancel;
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            // Načtení motorů z JSON souboru
            List<Motor> motorySeznam = Soubory.LoadJsonList<Motor>(CestaMotor);
            label2.Text = "Cesta = " + CestaMotor;
            Motor = [with(motorySeznam)];
            dataGridView1.DataSource = Motor;
            SetListBox();
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            // Načtení měničů z JSON souboru (opraveno z LoadFromCsv, protože CestaFM je .json)
            List<Menic> meniceSeznam = Soubory.LoadJsonList<Menic>(CestaFM);
            label2.Text = "Cesta = " + CestaFM;
            FM = [with(meniceSeznam)];
            SetListBox();
            dataGridView1.ClearSelection();
            dataGridView1.DataSource = FM;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            // Otevření složky s motory v Průzkumníku
            System.Diagnostics.Process.Start("explorer.exe", CestaMotor);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            // Otevření složky s měniči v Průzkumníku
            System.Diagnostics.Process.Start("explorer.exe", CestaFM);
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            // Otevření složky se stykači v Průzkumníku
            System.Diagnostics.Process.Start("explorer.exe", CestaKM);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            // Načtení stykačů z JSON souboru (opraveno z LoadFromCsv, protože CestaKM je .json)
            List<Stykac> stykaceSeznam = Soubory.LoadJsonList<Stykac>(CestaKM);
            label2.Text = "Cesta = " + CestaKM;
            KM = [with(stykaceSeznam)];
            SetListBox();
            dataGridView1.DataSource = KM;
        }

        private string SaveCesta { get; set; } = string.Empty;

        private BindingList<Stykac> KM = [];
        private readonly string CestaKM = Cesty.KM;

        private BindingList<Menic> FM = [];
        private readonly string CestaFM = Cesty.FM;

        private BindingList<Jistic> FA = [];
        private readonly string CestaJistic = Cesty.Jistic;

        private BindingList<Motor> Motor = [];
        private readonly string CestaMotor = Cesty.Motor;

        private void Button8_Click(object sender, EventArgs e)
        {
            // Uložení stykačů jako JSON a CSV
            Console.WriteLine($"Stykače uloženy jako Json a CSV");
            KM.ToList().SaveJsonList(Cesty.KM);
            KM.ToList().SaveToCsv(Cesty.KMCsv);
            dataGridView1.DataSource = null;
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            // Uložení měničů jako JSON a CSV
            Console.WriteLine($"Menice uloženy jako Json a CSV");
            FM.ToList().SaveJsonList(CestaFM);
            FM.ToList().SaveToCsv(Cesty.FMCsv);
            dataGridView1.DataSource = null;
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            // Uložení motorů jako JSON a CSV
            Console.WriteLine($"Motory uloženy jako Json a CSV");
            Motor.ToList().SaveJsonList(CestaMotor);
            Motor.ToList().SaveToCsv(Cesty.MotorCsv);
            dataGridView1.DataSource = null;
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            // Načtení jističů z JSON souboru
            List<Jistic> jisticeSeznam = Soubory.LoadJsonList<Jistic>(Cesty.Jistic);

            label2.Text = "Cesta = " + CestaJistic;
            FA = [with(jisticeSeznam)];
            SetListBox();
            dataGridView1.DataSource = FA;
        }

        private void Vytvořit_FormClosing(object sender, FormClosingEventArgs e)
        {
            Console.WriteLine($"Uloženy jako Json");
        }

        private void Button12_Click(object sender, EventArgs e)
        {
            // Uložení jističů jako JSON
            Console.WriteLine($"Jističe uloženy jako Json.");
            FA.ToList().SaveJsonList(Path.ChangeExtension(CestaJistic, ".json"));
            dataGridView1.DataSource = null;
        }
    }
}

