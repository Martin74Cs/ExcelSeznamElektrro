
using Aplikace.Sdilene;
using Aplikace.Upravy;
using Knihovna;
using Knihovna.Export;
using Knihovna.Tridy;
using Knihovna.Sdilene;
using System.Text;

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
            System.Diagnostics.Process.Start("explorer.exe", Informace.Instance.BasePath);
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
            Console.WriteLine("Informace.Instance.BasePath " + Informace.Instance.BasePath);
            string? Cesta = ZajistitSouborElektroJson();
            if (Cesta == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesta);
            if(Data.Count < 1) {
                Console.WriteLine("Soubor je prázdný " + Cesta);
            }

            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();

            if(result == DialogResult.OK) {
                Console.WriteLine($"Hotovo! Soubor JSON byl uložen do {Path.GetFileName(Cesta)}");
                Data.SaveJsonList(Cesta);
            }
        }

        /// <summary>
        /// Vlastní vývody mimo stroje
        /// </summary>
        private void Button11_Click(object sender, EventArgs e) {
            //var Vývody = Path.Combine(Cesty.Elektro, "Vývody.csv");
            //var Data = Soubory.LoadFromCsv<Zarizeni>(Vývody);

            var Vývody = Path.Combine(Cesty.VyvodyOstatniJson);
            if(!File.Exists(Vývody)) {
                Console.WriteLine("Soubor nebyl nalezen " + Vývody);
                Console.WriteLine("Soubor bude vytvořen!");
                //var prazdny = new List<Zarizeni>();
                Soubory.SaveJson(new List<Zarizeni>(), Cesty.VyvodyOstatniJson);
                Console.WriteLine("Znovu klikni na tlačítko. Soubor byl vytvořen!");
                return;
            }
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
            System.Diagnostics.Process.Start("explorer.exe", Informace.Instance.BasePath);
        }

        private void Button13_Click(object sender, EventArgs e) {
            string? cestaStroje = ZajistitSouborStrojeJson();
            if (cestaStroje == null) return;

            string? cestaElektro = ZajistitSouborElektroJson();
            if (cestaElektro == null) return;

            var Strojni = Soubory.LoadJsonList<Zarizeni>(cestaStroje);
            var Elektro = Soubory.LoadJsonList<Zarizeni>(cestaElektro);
            
            var table = new Shoda(Strojni, Elektro);
            var result = table.ShowDialog();
            if(result == DialogResult.OK) {
                Elektro.SaveJsonList(cestaElektro);
                Console.WriteLine($"Soubor {cestaElektro} -  aktualizován.");
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
        }

        private void Button14_Click(object sender, EventArgs e) {
            string? cestaStroje = ZajistitSouborStrojeJson();
            if (cestaStroje == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(cestaStroje);
            if(Data.Count > 0)
                Console.WriteLine($"Soubor {cestaStroje} -  načten.\npočet záznamů: {Data.Count}");
            else {
                Console.WriteLine($"Soubor je prázdný: {cestaStroje}");
                return;
            }
            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if(result == DialogResult.OK) {
                Data.SaveJsonList(cestaStroje);
            }
            else if(result == DialogResult.Cancel) {
                Console.WriteLine($"DialogResult.Cancel");
                Console.WriteLine($"Soubor : {Path.GetFileName(cestaStroje)} - ULOŽEN.");
                Data.SaveJsonList(cestaStroje);
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
            string? Cesta = ZajistitSouborElektroJson();
            if (Cesta == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesta);
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
                //var info = 
                Informace.Instance.BasePath = Folder.SelectedPath;
                //info.BasePath = Folder.SelectedPath;
                Console.WriteLine($"Složka nastavena na {Informace.Instance.BasePath}.");
                Informace.Instance.Ulozit();
            }

        }

        private void CestyToolStripMenuItem_Click(object sender, EventArgs e) {
            using var f = new WinForms.Nastaveni(); f.ShowDialog(this);
        }

        private void seznamToolStripMenuItem_Click(object sender, EventArgs e) {
            string? Cesta = ZajistitSouborElektroJson();
            if (Cesta == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesta);

            //vytvoření filtru, vlastnost, string, co dělat, negace
            var filtry = new List<FilterRule>
            {
                new(nameof(Zarizeni.Prikon), op: FilterOperator.IsNotNullOrEmpty),
                new(nameof(Zarizeni.Prikon), "-", op: FilterOperator.StartsWith,true),
                new(nameof(Zarizeni.Prikon), "-", op: FilterOperator.StartsWith,true),
            };
            var Vysledek = Soubory.ApplyFilter(Data, filtry).ToList();
            if(Vysledek.Count < 1) { Console.WriteLine("Nepsahuje data."); return; }

            string[] sloupceZarizeni = [
                nameof(Zarizeni.Tag),
                nameof(Zarizeni.Popis),
                nameof(Zarizeni.Prikon),
                nameof(Zarizeni.Napeti),
                nameof(Zarizeni.Menic),
                nameof(Zarizeni.BalenaJednotka),
                nameof(Zarizeni.Pozice)
            ];

            Vysledek.SaveToCsv(Path.ChangeExtension(Cesta, ".csv"), sloupceZarizeni);
            Vysledek.SaveXML(Path.ChangeExtension(Cesta, ".xml"), sloupceZarizeni);
            Vysledek.SaveHtmlStyle(Path.ChangeExtension(Cesta, ".html"), sloupceZarizeni);
            Vysledek.SavePdfGen(Path.ChangeExtension(Cesta, ".pdf"), null, sloupceZarizeni);
            Vysledek.SaveDocxGen(Path.ChangeExtension(Cesta, ".docx"), null, sloupceZarizeni);
        }

        private void kabelyToolStripMenuItem_Click(object sender, EventArgs e) {
            string? Cesta = ZajistitSouborElektroJson();
            if (Cesta == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesta);

            // Filtrování zařízení stejně jako u seznamu zařízení
            var filtry = new List<FilterRule>
            {
                new(nameof(Zarizeni.Prikon), op: FilterOperator.IsNotNullOrEmpty),
                new(nameof(Zarizeni.Prikon), "—", op: FilterOperator.StartsWith, true),
                new(nameof(Zarizeni.Prikon), "-", op: FilterOperator.StartsWith, true),
            };
            var filtrovanaData = Soubory.ApplyFilter(Data, filtry).ToList();

            var SeznamKabelu = new List<Trasa>();
            foreach (var z in filtrovanaData) {
                if (z.SeznamKabelu == null) continue;
                foreach (var k in z.SeznamKabelu) {
                    var kopieKabelu = new Trasa
                    {
                        Tag = z.Tag,
                        Rozvadec = k.Rozvadec,
                        RozvadecCislo = k.RozvadecCislo,
                        // Použijeme předznamenání tag ze zařízení, abychom zabránili shodným označením (WL1 -> P132WL1)
                        Oznaceni = (z.Tag ?? "").Replace(" ", "") + (k.Oznaceni ?? "").Replace(" ", ""),
                        Kabel = k.Kabel,
                        PocetZil = k.PocetZil,
                        Prurezmm2 = k.Prurezmm2,
                        PrurezFt = k.PrurezFt,
                        Druh = k.Druh,
                        OdkudSvokra = k.OdkudSvokra,
                        Mezera = k.Mezera,
                        Patro = k.Patro,
                        Predmet = k.Predmet,
                        Svorka = k.Svorka,
                        Delka = k.Delka,
                        Popis = k.Popis
                    };
                    SeznamKabelu.Add(kopieKabelu);
                }
            }

            string directory = Path.GetDirectoryName(Cesta)!;
            string targetBase = Path.Combine(directory, "Elektro.Kabely");

            Console.WriteLine($"Generování seznamu kabelů do {targetBase}.*");
            
            string[] sloupceKabelu = [
                nameof(Trasa.Tag),
                //nameof(Trasa.Rozvadec),
                //nameof(Trasa.RozvadecCislo),
                nameof(Trasa.RozvadecAll),
                nameof(Trasa.Oznaceni),
                //nameof(Trasa.Kabel),
                //nameof(Trasa.PocetZil),
                //nameof(Trasa.Prurezmm2),
                nameof(Trasa.KabelAll),
                //nameof(Trasa.Druh),
                nameof(Trasa.Delka),
                nameof(Trasa.Popis)
            ];

            SeznamKabelu.SaveToCsv(targetBase + ".csv", sloupceKabelu);
            SeznamKabelu.SaveXML(targetBase + ".xml", sloupceKabelu);
            SeznamKabelu.SaveHtmlStyle(targetBase + ".html", sloupceKabelu);
            SeznamKabelu.SavePdfGen(targetBase + ".pdf", null, sloupceKabelu);
            SeznamKabelu.SaveDocxGen(targetBase + ".docx", null, sloupceKabelu);
        }

        /// <summary>
        /// Zajistí, že existuje soubor Elektro.Data.json.
        /// Pokud neexistuje, nabídne uživateli dialog pro jeho výběr nebo možnost zkopírovat jej ze souboru Strojni.
        /// </summary>
        /// <returns>Cesta k souboru, nebo null, pokud se soubor nepodařilo zajistit.</returns>
        private string? ZajistitSouborElektroJson() {
            var Cesta = Informace.Instance.SouborElektroJson;
            if(!File.Exists(Cesta)) { 
                Cesta = Soubory.ShowOpenFileDialog("Json soubor (*.json)|*.json", Informace.Instance.BasePath);
                if(string.IsNullOrEmpty(Cesta) || !File.Exists(Cesta)) {
                    Console.WriteLine("Výběr souboru Elektro byl stornován nebo soubor neexistuje.");
                    
                    // Nabídneme vytvoření kopie ze souboru Strojni
                    if(MessageBox.Show("Chcete vytvořit kopii souboru ze souboru Strojni?", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK) {
                        string? cestaStroje = ZajistitSouborStrojeJson();
                        if (cestaStroje == null) {
                            Console.WriteLine("Nelze vytvořit kopii, protože chybí zdrojový soubor Strojni.");
                            return null;
                        }
                        
                        Informace.Instance.SouborElektroJson = Path.Combine(Informace.Instance.BasePath, "Elektro.Data.json");
                        Informace.Instance.Ulozit();
                        Cesta = Informace.Instance.SouborElektroJson;
                        
                        try {
                            File.Copy(cestaStroje, Cesta, overwrite: true);
                            if(File.Exists(Cesta)) {
                                Console.WriteLine($"Soubor {Cesta} - zkopírován ze {cestaStroje}.");
                            }
                        }
                        catch (Exception ex) {
                            Console.WriteLine($"Chyba při kopírování souboru: {ex.Message}");
                            return null;
                        }
                    }
                    else {
                        return null;
                    }
                }
                else {
                    Informace.Instance.SouborElektroJson = Cesta;
                    Informace.Instance.Ulozit();
                }
            }
            return Cesta;
        }

        /// <summary>
        /// Zajistí, že existuje soubor se stroji.
        /// Pokud neexistuje, nabídne uživateli dialog pro jeho výběr.
        /// </summary>
        /// <returns>Cesta k souboru, nebo null, pokud se soubor nepodařilo zajistit.</returns>
        private string? ZajistitSouborStrojeJson() {
            var Cesta = Informace.Instance.SouborStrojeJson;
            if(!File.Exists(Cesta)) {
                Cesta = Soubory.ShowOpenFileDialog("Json soubor (*.json)|*.json", Informace.Instance.BasePath);
                if(string.IsNullOrEmpty(Cesta) || !File.Exists(Cesta)) {
                    Console.WriteLine("Výběr souboru strojů byl stornován nebo soubor neexistuje.");
                    return null;
                }
                else {
                    Informace.Instance.SouborStrojeJson = Cesta;
                    Informace.Instance.Ulozit();
                }
            }
            return Cesta;
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
