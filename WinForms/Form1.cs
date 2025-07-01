using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Tridy;
using Aplikace.Upravy;
using System.ComponentModel;
using System.Globalization;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;
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
            //Převod->json,csv
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

        //private async void Button5_Click(object sender, EventArgs e)
        //{
        //    await Task.Run(() => LigthChem.AddKabely());
        //}

        private void Button8_Click(object sender, EventArgs e)
        {
            //string cestaData = Path.Combine(Cesty.Elektro, @"ElektroData.csv");
            System.Diagnostics.Process.Start("explorer.exe", Cesty.Elektro);
        }

        private async void Button9_Click(object sender, EventArgs e)
        {
            //await Task.Run(() => LigthChem.AddVyvody());
            await Task.Run(() => LigthChem.Rozvadec());
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
        private void MístnostiToolStripMenuItem1_Click(object sender, EventArgs e)
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


        private void Button10_Click(object sender, EventArgs e)
        {
            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (Data.Count < 1) Data.Add(new Zarizeni());
                Data.SaveJsonList(Cesty.ElektroDataJson);
                if(MessageBox.Show("Aktualiyace CSV", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK) { 
                    Data.SaveToCsv(Cesty.ElektroDataCsv);
                    Data.SaveXML(Path.ChangeExtension(Cesty.ElektroDataCsv, ".xml"));
                    Data.SaveHtml(Path.ChangeExtension(Cesty.ElektroDataCsv, ".html"));
                }
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            //var Vývody = Path.Combine(Cesty.Elektro, "Vývody.csv");
            //var Data = Soubory.LoadFromCsv<Zarizeni>(Vývody);

            var Vývody = Path.Combine(Cesty.Elektro, "Vývody.json");
            if (!File.Exists(Vývody)) { Console.WriteLine("Soubor nebyl nalezen " + Vývody); return; }
            var Data = Soubory.LoadJsonList<Zarizeni>(Vývody);

            //var DataBind = new BindingList<Zarizeni>(Data);
            var table = new Table(Data);
            SkrytSloupce(table.dataGridView1);
            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                //přidat prázdný záznam
                if (Data.Count < 1) Data.Add(new Zarizeni());

                //Data.SaveToCsv(Vývody);
                Data.SaveJsonList(Vývody);

                //if (MessageBox.Show("Aktualiyace CSV", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK)
                //    Data.SaveToCsv(Cesty.ElektroDataCsv);
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        private static void SkrytSloupce(DataGridView data)
        {
            //skryje sloupce, které nechceme zobrazit
            data.Columns["Patro"]?.Visible = false;
            data.Columns["HP"]?.Visible = false;
            data.Columns["Delka"]?.Visible = false;
            data.Columns["IsExist"]?.Visible = false;
            data.Columns["IsExistElektro"]?.Visible = false;
            data.Columns["Bod"]?.Visible = false;
            data.Columns["BodElektro"]?.Visible = false;
            data.Columns["PID"]?.Visible = false;
            data.Columns["Pocet"]?.Visible = false;
            data.Columns["Radek"]?.Visible = false;
            data.Columns["Id"]?.Visible = false;
            data.Columns["Otoceni"]?.Visible = false;

            data.Columns["Nic"]?.Visible = false;
            data.Columns["AWG"]?.Visible = false;
            data.Columns["Delkaft"]?.Visible = false;

            //data.Columns["PrurezMM2"].Visible = false;
            //data.Columns["Rozvadec"].Visible = false;
            //data.Columns["RozvadecCislo"].Visible = false;
            //data.Columns["RozvadecOznačení"].Visible = false;
            //data.Columns["Kabel"].Visible = false;
            //data.Columns["Motor"].Visible = false;

            //data.Columns["Vykres"].Visible = false;
            //data.Columns["Vodice"].Visible = false;
            //data.Columns["Motor"].Visible = false;
        }


        private void Button12_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", Cesty.Elektro);
        }

        private void Button13_Click(object sender, EventArgs e)
        {
            //string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string cesta1 = Path.Combine(Cesty.Elektro, "Pid", @"UpravaZnovu.006.json");
            var Strojni = Soubory.LoadJsonList<Zarizeni>(Path.ChangeExtension(cesta1, ".json"));

            var Elektro = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            var table = new Shoda(Strojni, Elektro);
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
                Elektro.SaveJsonList(Cesty.ElektroDataJson);
            }

            //foreach (var itemEl in Elektro.ToHashSet())
            //{
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
            Elektro.SaveJsonList(Cesty.ElektroDataJson);
        }

        private void Button14_Click(object sender, EventArgs e)
        {
            //string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string cesta1 = Path.Combine(Cesty.Elektro, "Pid", @"UpravaZnovu.006.json");
            var Data = Soubory.LoadJsonList<Zarizeni>(Path.ChangeExtension(cesta1, ".json"));

            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (Data.Count < 1) Data.Add(new Zarizeni());
                Data.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));

                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        private async void ExpotrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Převod extrahovaných dat z Dwg do Xls s následným převodem do Json
            await Task.Run(() => LigthChem.DwgXlsToJsonCsv());
        }

        private void PropojeniToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var table = new Rozvaděč();
            //var result = table.ShowDialog();
            table.ShowDialog();
        }

        //Vývody stavba
        private void Button5_Click(object sender, EventArgs e)
        {
            var Vývody = Path.Combine(Cesty.Elektro, "Vývody.Stavba.json");
            var Data = Soubory.LoadJsonList<Zarizeni>(Vývody);

            var table = new Table(Data);
            SkrytSloupce(table.dataGridView1);
            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                //přidat prázdný záznam
                if (Data.Count < 1) Data.Add(new Zarizeni());
                Data.SaveJsonList(Vývody);
            }
        }

        private void PříkonCelkemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            Console.WriteLine($"Příkon celkem: {Data.Sum(x => double.TryParse(x.Prikon, out var p ) ? p : 0.0) } W");
            Console.WriteLine($"Příkon FAZE 1: {Data.Where(x => x.Etapa == "FAZE 1").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");
            Console.WriteLine($"Příkon FAZE 2: {Data.Where(x => x.Etapa == "FAZE 2").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");

            var Topeni = Data.Where(x => x.RozvadecOznačení != "RT01");
            Console.WriteLine($"Příkon bez topení");
            Console.WriteLine($"Příkon celkem: {Topeni.Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} W");
            Console.WriteLine($"Příkon FAZE 1: {Topeni.Where(x => x.Etapa == "FAZE 1").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");
            Console.WriteLine($"Příkon FAZE 2: {Topeni.Where(x => x.Etapa == "FAZE 2").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");
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
