using Aplikace.Sdilene;
using Aplikace.Tridy;
using Aplikace.Upravy;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;

namespace WinForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Převod stroju na JSON a CSV
        private async void Button2_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.StrojniToJsonCsv());
            //Console.SetOut(new ListBoxWriter(listBox1));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.SetOut(new ListBoxWriter(listBox1));
        }


        private void ListBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            var box = (ListBox)sender;
            textBox1.Text = box.Text;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private async void Button3_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.DoplneniCsvToJson());
        }

        private async void Button4_Click(object sender, EventArgs e)
        {
            await Task.Run(() => Soubory.KillExcel());
        }

        private async void Button5_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.AddKabely());
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            string cestaData = Path.Combine(Cesty.Elektro, @"ElektroData.csv");
            System.Diagnostics.Process.Start("explorer.exe", cestaData);
        }

        private async void Button9_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.AddVyvody());
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string cesta = Cesty.BasePath;
            System.Diagnostics.Process.Start("explorer.exe", cesta);
        }

        private void SeznamyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var vyvorit = new Vytvořit();

            SetTable(vyvorit);

            // Zobrazíme druhý formulář jako modální dialog
            var result = vyvorit.ShowDialog();
            if (result == DialogResult.OK)
            {
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        /// <summary>Nastavení pomocného okna </summary>
        private void SetTable(Form form)
        {
            // Vypočteme střed Form1 a posuneme Form2 tam
            int x = this.Location.X + (this.Width - form.Width) / 2;
            int y = this.Location.Y + (this.Height - form.Height) / 2;

            // Nastavíme pozici druhého formuláře
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(x, y);
        }

        /// <summary> Místnosti - otevřít seznam </summary>
        private async void MístnostiToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //použitá cesta z Místnosti.cs          
            System.Diagnostics.Process.Start("explorer.exe", Cesty.MistnostiXLs);
        }

        /// <summary> Místnosti - vytvoření seznamu </summary>
        private async void GenerovatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await Task.Run(() => Místnosti.VytvoritSeznamy());
        }

        private async void Button6_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.JsonToExcel());
        }

        private async void Button7_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.AddProud());
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            var table = new Table(Data);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                Data.SaveJsonList(Cesty.ElektroDataJson);
                if (MessageBox.Show("Aktualiyace CSV", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    Data.SaveToCsv(Cesty.ElektroDataCsv);
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            var Vývody = Path.Combine(Cesty.Elektro, "Vývody.csv");
            var Data = Soubory.LoadFromCsv<Zarizeni>(Vývody);
            //var DataBind = new BindingList<Zarizeni>(Data);
            var table = new Table(Data);
            //table.Initialize();

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                //přidat prázdný záznam
                if (Data.Count < 1) Data.Add(new Zarizeni());
                Data.SaveToCsv(Vývody);
                //if (MessageBox.Show("Aktualiyace CSV", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK)
                //    Data.SaveToCsv(Cesty.ElektroDataCsv);
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }


        private void Button12_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", Cesty.Elektro);
        }
        
    }

    public class ListBoxWriter(ListBox listBox) : TextWriter
    {
        private readonly ListBox _listBox = listBox;
        private readonly SynchronizationContext _context = SynchronizationContext.Current;

        public override Encoding Encoding => Encoding.UTF8;

        public override void WriteLine(string value)
        {
            //_context.Post(_ => _listBox.Items.Add(value), null);
            _context.Post(_ =>
            {
                _listBox.Items.Add(value);
                _listBox.TopIndex = _listBox.Items.Count - 1; // ← automatické scrollování dolů
            }, null);
        }

        //public override void Write(char value)
        //{
        //    // Nepřepisujeme po znacích, pouze řádky (volitelné)
        //}
    }
}
