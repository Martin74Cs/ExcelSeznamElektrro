using Aplikace.Excel;
using Aplikace.Export;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace WinForms
{
    public partial class Form1: Form {
        public Form1() {
            InitializeComponent();
        }

        //Převod stroju na JSON a CSV z xls.
        //Xls je podklad strjů a zařízení z projektu strojní
        private async void Button2_Click(object sender, EventArgs e) {
            //Převod->json,csv
            await Task.Run(() => LigthChem.StrojniToJsonCsv());
            //Console.SetOut(new ListBoxWriter(listBox1));
        }

        private void Form1_Load(object sender, EventArgs e) {
            Console.SetOut(new ListBoxWriter(listBox1));

        }


        private void ListBox1_SelectedIndexChanged_1(object sender, EventArgs e) {
            var box = (ListBox)sender;
            textBox1.Text = box.Text;
        }

        private void Button1_Click(object sender, EventArgs e) {
            Close();
        }

        private async void Button3_Click(object sender, EventArgs e) {
            await Task.Run(() => LigthChem.DoplneniCsvToJson());
        }

        private async void Button4_Click(object sender, EventArgs e) {
            await Task.Run(() => Soubory.KillExcel());
        }

        //private async void Button5_Click(object sender, EventArgs e)
        //{
        //    await Task.Run(() => LigthChem.AddKabely());
        //}

        private void Button8_Click(object sender, EventArgs e) {
            //string cestaData = Path.Combine(Cesty.Elektro, @"ElektroData.csv");
            System.Diagnostics.Process.Start("explorer.exe", Informace.Create.BasePath);
        }

        private async void Button9_Click(object sender, EventArgs e) {
            //await Task.Run(() => LigthChem.AddVyvody());
            await Task.Run(() => LigthChem.Rozvadec());
        }

        //Otevřít složku projektu v Průzkumníku
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e) {
            string cesta = Informace.Adresar;
            System.Diagnostics.Process.Start("explorer.exe", cesta);
        }

        private void SeznamyToolStripMenuItem_Click(object sender, EventArgs e) {
            var vyvorit = new Vytvořit();

            SetTable(vyvorit);

            // Zobrazíme druhý formulář jako modální dialog
            var result = vyvorit.ShowDialog();
            if(result == DialogResult.OK) {
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        /// <summary>Nastavení pomocného okna </summary>
        private void SetTable(Form form) {
            // Vypočteme střed Form1 a posuneme Form2 tam
            int x = this.Location.X + (this.Width - form.Width) / 2;
            int y = this.Location.Y + (this.Height - form.Height) / 2;

            // Nastavíme pozici druhého formuláře
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(x, y);
        }

        /// <summary> Místnosti - otevřít seznam </summary>
        private void MístnostiToolStripMenuItem1_Click(object sender, EventArgs e) {
            //použitá cesta z Místnosti.cs          
            System.Diagnostics.Process.Start("explorer.exe", Cesty.MistnostiXLs);
        }

        /// <summary> Místnosti - vytvoření seznamu </summary>
        private async void GenerovatToolStripMenuItem_Click(object sender, EventArgs e) {
            await Task.Run(() => Místnosti.VytvoritSeznamy());
        }

        private async void Button6_Click(object sender, EventArgs e) {
            await Task.Run(() => LigthChem.JsonToExcel());
        }


        private void Button10_Click(object sender, EventArgs e) {
            var Data = Soubory.LoadJsonList<Zarizeni>(Informace.Create.SouborElektroJson);
            if(Data.Count < 1) {
                Console.WriteLine("Soubor je prázdný " + Informace.Create.SouborElektroJson);
                if(MessageBox.Show("Kopie souborů ze souboru Strojni", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK) {

                    var Adresar = Path.GetDirectoryName(Informace.Create.SouborStrojeJson);
                    using var info = Informace.Create;
                    info.SouborElektroJson = Path.Combine(Adresar,"Elektro.Data.json");

                    if(!File.Exists(Informace.Create.SouborStrojeJson)) {
                        Console.WriteLine("Soubor nebyl nalezen " + Informace.Create.SouborStrojeJson); return;
                    }
                    File.Copy(Informace.Create.SouborStrojeJson, Informace.Create.SouborElektroJson);
                    if(File.Exists(Informace.Create.SouborElektroJson))
                        Console.WriteLine($"Soubor {Informace.Create.SouborElektroJson} -  zkopírován ze {Informace.Create.SouborStrojeJson}.");
                    //nové načtení bylo prázdné, takže načteme znovu po kopírování
                    Data = Soubory.LoadJsonList<Zarizeni>(Informace.Create.SouborElektroJson);
                }
            }

            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();

            if(result == DialogResult.OK) {
                //if (Data.Count < 1) Data.Add(new Zarizeni());

                //Soubou znovu uložit je možné že nastaly změny v souboru
                Data.SaveJsonList(Informace.Create.SouborElektroJson);

                if(MessageBox.Show("Aktualizace CSV", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK) {
                    Data.SaveToCsv(Path.ChangeExtension(Informace.Create.SouborElektroJson, ".csv"));
                    Data.SaveXML(Path.ChangeExtension(Informace.Create.SouborElektroJson, ".xml"));
                    Data.SaveHtmlStyle(Path.ChangeExtension(Informace.Create.SouborElektroJson, ".html"));
                    //Data.SaveDocx(Path.ChangeExtension(Informace.Create.SouborElektroJson, ".docx"));
                    Data.SavePdfGen(Path.ChangeExtension(Informace.Create.SouborElektroJson, ".pdf"));
                    Data.SaveDocxGen(Path.ChangeExtension(Informace.Create.SouborElektroJson, ".docx"));
                }
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        /// <summary>
        /// Vlastní vývody mimo stroje
        /// </summary>
        private void Button11_Click(object sender, EventArgs e) {
            //var Vývody = Path.Combine(Cesty.Elektro, "Vývody.csv");
            //var Data = Soubory.LoadFromCsv<Zarizeni>(Vývody);

            var Vývody = Path.Combine(Cesty.VyvodyJson);
            if(!File.Exists(Vývody)) { Console.WriteLine("Soubor nebyl nalezen " + Vývody); return; }
            var Data = Soubory.LoadJsonList<Zarizeni>(Vývody);

            //var DataBind = new BindingList<Zarizeni>(Data);
            var table = new Table(Data);
            SkrytSloupce(table.dataGridView1);
            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if(result == DialogResult.OK) {
                //přidat prázdný záznam
                if(Data.Count < 1) Data.Add(new Zarizeni());

                //Data.SaveToCsv(Vývody);
                Data.SaveJsonList(Vývody);

                //if (MessageBox.Show("Aktualiyace CSV", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK)
                //    Data.SaveToCsv(Cesty.ElektroDataCsv);
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        private static void SkrytSloupce(DataGridView data) {
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

        /// <summary> Průzkumník tedy složka projektu </summary>
        private void Button12_Click(object sender, EventArgs e) {
            System.Diagnostics.Process.Start("explorer.exe", Informace.Create.BasePath);
        }

        //Doplmění dat do Elektro z aktualizovaného Strojního seznamu,
        //kontrola shod podle Tagu, pokud je shoda jedna, doplní se pouze prázdné bunky, pokud je více shod, vypíše se počet shod do popisu pro kontrolu
        private void Button13_Click(object sender, EventArgs e) {
            //string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            //string cesta1 = Path.Combine(Cesty.Elektro, "Pid", @"UpravaZnovu.006.json");

            //var Strojni = Soubory.LoadJsonList<Zarizeni>(Path.ChangeExtension(cesta1, ".json"));
            var Strojni = Soubory.LoadJsonList<Zarizeni>(Informace.Create.SouborStrojeJson);

            var Elektro = Soubory.LoadJsonList<Zarizeni>(Informace.Create.SouborElektroJson);
            var table = new Shoda(Strojni, Elektro);
            var result = table.ShowDialog();
            if(result == DialogResult.OK) {
                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
                Elektro.SaveJsonList(Informace.Create.SouborElektroJson);
                Console.WriteLine($"Soubor {Informace.Create.SouborElektroJson} -  aktualizován.");
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
            //Elektro.SaveJsonList(Cesty.ElektroDataJson);
        }

        //Otevřít Json seznamu strojů a zařízení, jen kontrola převodu XLS na Json
        private void Button14_Click(object sender, EventArgs e) {
            //var cesta1 = Informace.Create.SouborStrojeJson;
            if(!File.Exists(Informace.Create.SouborStrojeJson)) {
                var cesta1 = Soubory.ShowOpenFileDialog("Json soubor (*.json)|*.json");
                if(string.IsNullOrEmpty(cesta1) || !File.Exists(cesta1)) {
                    Console.WriteLine("Výběr souboru byl stornován nebo soubor neexistuje."); return;
                }
                else {
                    //uložíme cestu do singletonu pro další použití pokud tam již nebyla
                    using var Cesta = Informace.Create;
                    Cesta.SouborStrojeJson = cesta1;
                }
            }
            //var Data = Soubory.LoadJsonList<Zarizeni>(Path.ChangeExtension(cesta1, ".json"));

            var Data = Soubory.LoadJsonList<Zarizeni>(Informace.Create.SouborStrojeJson);
            if(Data.Count > 0)
                Console.WriteLine($"Soubor {Informace.Create.SouborStrojeJson} -  načten.\npočet záznamů: {Data.Count}");
            else {
                Console.WriteLine($"Soubor je prázdný: {Informace.Create.SouborStrojeJson}");
                return;
            }
            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if(result == DialogResult.OK) {

                Data.SaveJsonList(Informace.Create.SouborStrojeJson);

                //if (Data.Count < 1) Data.Add(new Zarizeni());
                //Data.SaveJsonList(Informace.Create.SouborStrojeJson);

                // Zde můžete provést další akce po zavření dialogu
                // Například načíst data nebo aktualizovat UI
            }
        }

        private async void ExpotrToolStripMenuItem_Click(object sender, EventArgs e) {
            //Převod extrahovaných dat z Dwg do Xls s následným převodem do Json
            await Task.Run(() => LigthChem.DwgXlsToJsonCsv());
        }

        private void PropojeniToolStripMenuItem_Click(object sender, EventArgs e) {
            var table = new Rozvaděč();
            //var result = table.ShowDialog();
            table.ShowDialog();
        }

        //Vývody stavba
        private void Button5_Click(object sender, EventArgs e) {

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.VyvodyStavbaJson);

            var table = new Table(Data);
            SkrytSloupce(table.dataGridView1);
            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if(result == DialogResult.OK) {
                //přidat prázdný záznam
                if(Data.Count < 1) Data.Add(new Zarizeni());
                Data.SaveJsonList(Cesty.VyvodyStavbaJson);
            }
        }

        private void PříkonCelkemToolStripMenuItem_Click(object sender, EventArgs e) {
            var Data = Soubory.LoadJsonList<Zarizeni>(Informace.Create.SouborElektroJson);
            Console.WriteLine($"Příkon celkem: {Data.Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} W");
            Console.WriteLine($"Příkon FAZE 1: {Data.Where(x => x.Etapa == "FAZE 1").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");
            Console.WriteLine($"Příkon FAZE 2: {Data.Where(x => x.Etapa == "FAZE 2").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");

            var Topeni = Data.Where(x => x.RozvadecOznačení != "RT01");
            Console.WriteLine($"Příkon bez topení");
            Console.WriteLine($"Příkon celkem: {Topeni.Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} W");
            Console.WriteLine($"Příkon FAZE 1: {Topeni.Where(x => x.Etapa == "FAZE 1").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");
            Console.WriteLine($"Příkon FAZE 2: {Topeni.Where(x => x.Etapa == "FAZE 2").Sum(x => double.TryParse(x.Prikon, out var p) ? p : 0.0)} kW");
        }

        private void NastavSložkuProjektuToolStripMenuItem_Click(object sender, EventArgs e) {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            //string file = Path.Combine(appData, "ElektroData", "data.txt");
            string file = Path.Combine(appData, "Elektro", "data.txt");
            Directory.CreateDirectory(Path.GetDirectoryName(file)!);

            //File.WriteAllText(file, @"C:\ElektroData");
            //OpenFileDialog openFileDialog = new OpenFileDialog
            //{
            //    InitialDirectory = @"C:\ElektroData",
            //    Title = "Vyberte složku projektu",
            //    CheckFileExists = false,
            //    CheckPathExists = true,
            //    //FileName = "Vyberte složku projektu"
            //};
            //var dialog = openFileDialog.ShowDialog();

            FolderBrowserDialog Folder = new() {
                Description = "Vyber složku s projektem",
                UseDescriptionForTitle = true // .NET 6+ moderní styl
            };
            if(Folder.ShowDialog() == DialogResult.OK) {
                using var info = Informace.Create;
                info.BasePath = Folder.SelectedPath;
                Console.WriteLine($"Složka nastavena na {info.BasePath}.");
            }

        }

        private void CestyToolStripMenuItem_Click(object sender, EventArgs e) {
            using var f = new Aplikace.Forms.Nastaveni(); f.ShowDialog(this);
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
