
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Upravy;
using ExcelGenerateSeznam;
using Knihovna;
using Knihovna.Export;
using Knihovna.KabelyXls;
using Knihovna.Sdilene;
using Knihovna.Shared.Tridy;
using Knihovna.Tridy;
using System.Text;

namespace WinForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent(); //.WaitAsync(cancellation);
        }

        //Převod stroju na JSON a CSV z xls.
        //Xls je podklad strjů a zařízení z projektu strojní
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
            System.Diagnostics.Process.Start("explorer.exe", Informace.Instance.BasePath);
        }

        private async void Button9_Click(object sender, EventArgs e)
        {
            //await Task.Run(() => LigthChem.AddVyvody());
            await Task.Run(() => LigthChem.Rozvadec());
        }

        //Otevřít složku projektu v Průzkumníku
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string cesta = Informace.Adresar;
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
            this.Hide();
            Console.WriteLine("Informace.Instance.BasePath " + Informace.Instance.BasePath);
            string? Cesta = ZajistitSouborElektroJson();
            if (Cesta == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesta);
            if (Data.Count < 1)
            {
                Console.WriteLine("Soubor je prázdný " + Cesta);
            }

            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();

            if (result == DialogResult.OK)
            {
                Console.WriteLine($"Hotovo! Soubor JSON byl uložen do {Path.GetFileName(Cesta)}");
                Data.SaveJsonList(Cesta);
            }
            this.Show();
        }

        /// <summary>
        /// Vlastní vývody mimo stroje
        /// </summary>
        private void Button11_Click(object sender, EventArgs e)
        {
            //var Vývody = Path.Combine(Cesty.Elektro, "Vývody.csv");
            //var Data = Soubory.LoadFromCsv<Zarizeni>(Vývody);

            var Vývody = Cesty.VyvodyOstatniJson;
            if (!File.Exists(Vývody))
            {
                Console.WriteLine("Soubor nebyl nalezen " + Vývody);
                Console.WriteLine("Soubor bude vytvořen!");
                //var prazdny = new List<Zarizeni>();
                Soubory.SaveJson(new List<Zarizeni>(), Vývody);
                Console.WriteLine("Znovu klikni na tlačítko. Soubor byl vytvořen!");
                return;
            }
            var Data = Soubory.LoadJsonList<Zarizeni>(Vývody);
            Console.WriteLine($"Soubor {Vývody} - načten.");

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

        /// <summary> Průzkumník tedy složka projektu </summary>
        private void Button12_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", Informace.Instance.BasePath);
        }

        private void Button13_Click(object sender, EventArgs e)
        {
            string? cestaStroje = ZajistitSouborStrojeJson();
            if (cestaStroje == null) return;

            string? cestaElektro = ZajistitSouborElektroJson();
            if (cestaElektro == null) return;

            var Strojni = Soubory.LoadJsonList<Zarizeni>(cestaStroje);
            var Elektro = Soubory.LoadJsonList<Zarizeni>(cestaElektro);

            var table = new Shoda(Strojni, Elektro);
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
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

        private void Button14_Click(object sender, EventArgs e)
        {
            string? cestaStroje = ZajistitSouborStrojeJson();
            if (cestaStroje == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(cestaStroje);
            if (Data.Count > 0)
                Console.WriteLine($"Soubor {cestaStroje} -  načten.\npočet záznamů: {Data.Count}");
            else
            {
                Console.WriteLine($"Soubor je prázdný: {cestaStroje}");
                return;
            }
            var table = new Table(Data);
            SetTable(table);

            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                Data.SaveJsonList(cestaStroje);
            }
            else if (result == DialogResult.Cancel)
            {
                Console.WriteLine($"DialogResult.Cancel");
                Console.WriteLine($"Soubor : {Path.GetFileName(cestaStroje)} - ULOŽEN.");
                Data.SaveJsonList(cestaStroje);
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

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.VyvodyStavbaJson);

            var table = new Table(Data);
            SkrytSloupce(table.dataGridView1);
            // Zobrazíme druhý formulář jako modální dialog
            var result = table.ShowDialog();
            if (result == DialogResult.OK)
            {
                //přidat prázdný záznam
                if (Data.Count < 1) Data.Add(new Zarizeni());
                Data.SaveJsonList(Cesty.VyvodyStavbaJson);
            }
        }

        private void PříkonCelkemToolStripMenuItem_Click(object sender, EventArgs e)
        {
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

        private void NastavSložkuProjektuToolStripMenuItem_Click(object sender, EventArgs e)
        {
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

            FolderBrowserDialog Folder = new()
            {
                Description = "Vyber složku s projektem",
                UseDescriptionForTitle = true // .NET 6+ moderní styl
            };
            if (Folder.ShowDialog() == DialogResult.OK)
            {
                //var info = 
                Informace.Instance.BasePath = Folder.SelectedPath;
                //info.BasePath = Folder.SelectedPath;
                Console.WriteLine($"Složka nastavena na {Informace.Instance.BasePath}.");
                Informace.Instance.Ulozit();
            }

        }

        private void CestyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var f = new WinForms.Nastaveni(); f.ShowDialog(this);
        }

        private void KontrolaCestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var f = new WinForms.FormCesty(); f.ShowDialog(this);
        }

        private void SeznamToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string? Cesta = ZajistitSouborElektroJson();
            if (Cesta == null) return;

            var Vysledek = Soubory.LoadJsonList<Zarizeni>(Cesta);

            //Už neplatí
            //vytvoření filtru, vlastnost, string, co dělat, negace
            //var filtry = new List<FilterRule>
            //{
            //    new(nameof(Zarizeni.Prikon), op: FilterOperator.IsNotNullOrEmpty),
            //    new(nameof(Zarizeni.Prikon), "-", op: FilterOperator.StartsWith,true),
            //    new(nameof(Zarizeni.Prikon), "-", op: FilterOperator.StartsWith,true),
            //};
            //var Vysledek = Soubory.ApplyFilter(Data, filtry).ToList();
            //if (Vysledek.Count < 1) { Console.WriteLine("Data nebyla nalezena."); return; }

            string[] sloupceZarizeni = [
                nameof(Zarizeni.Tag),
                nameof(Zarizeni.Popis),
                nameof(Zarizeni.Prikon),
                nameof(Zarizeni.Napeti),
                nameof(Zarizeni.Menic),
                nameof(Zarizeni.BalenaJednotka),
                nameof(Zarizeni.Pozice)
            ];

            var Adresar = Path.GetDirectoryName(Cesta);
            string targetBase = Path.Combine(Adresar, "Výstup");
            if (!Directory.Exists(targetBase)) { Directory.CreateDirectory(targetBase); }
            Cesta = Path.Combine(targetBase, Path.GetFileName(Cesta));

            Vysledek.SaveToCsv(Path.ChangeExtension(Cesta, ".csv"), sloupceZarizeni);
            Vysledek.SaveXML(Path.ChangeExtension(Cesta, ".xml"), sloupceZarizeni);
            Vysledek.SaveHtmlStyle(Path.ChangeExtension(Cesta, ".html"),"Seznam zařízení", sloupceZarizeni);
            Vysledek.SavePdfGen(Path.ChangeExtension(Cesta, ".pdf"), "Seznam zařízení", sloupceZarizeni);
            Vysledek.SaveDocxGen(Path.ChangeExtension(Cesta, ".docx"), "Seznam zařízení", sloupceZarizeni);
            Vysledek.SaveXlsxGen(Path.ChangeExtension(Cesta, ".xlsx"), "Seznam zařízení", sloupceZarizeni);
        }

        private void KabelyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string? Cesta = ZajistitSouborElektroJson();
            if (Cesta == null) return;

            var Data = Soubory.LoadJsonList<Zarizeni>(Cesta);

            //Už neplatí
            // Filtrování zařízení stejně jako u seznamu zařízení
            //var filtry = new List<FilterRule>
            //{
            //    new(nameof(Zarizeni.Prikon), op: FilterOperator.IsNotNullOrEmpty),
            //    new(nameof(Zarizeni.Prikon), "—", op: FilterOperator.StartsWith, true),
            //    new(nameof(Zarizeni.Prikon), "-", op: FilterOperator.StartsWith, true),
            //};
            //var filtrovanaData = Soubory.ApplyFilter(Data, filtry).ToList();

            //Vyvoření seznamu
            var SeznamKabelu = new List<Trasa>();
            foreach (var z in Data)
            {
                if (z.SeznamKabelu == null) continue;
                foreach (var k in z.SeznamKabelu)
                {
                    var kopieKabelu = new Trasa
                    {
                        Tag = z.Tag,
                        Rozvadec = k.Rozvadec,
                        RozvadecCislo = k.RozvadecCislo,
                        // Použijeme předznamenání tag ze zařízení, abychom zabránili shodným označením (WL1 -> P132-WL1)
                        Oznaceni = (z.Tag ?? "").Replace(" ", "") + "-"+ (k.Oznaceni ?? "").Replace(" ", ""),
                        Kabel = k.Kabel,
                        KabelVelikost = k.PocetZil + "x" + k.Prurezmm2,
                        //PocetZil = k.PocetZil,
                        //Prurezmm2 = k.Prurezmm2,
                        PrurezFt = k.PrurezFt,
                        Druh = k.Tag + "-" + k.Popis,
                        OdkudSvorka = string.IsNullOrEmpty(k.OdkudSvorka) ? "SZ" : k.OdkudSvorka, 
                        Mezera = k.Mezera,
                        Patro = k.Patro,
                        Predmet = k.Predmet,
                        //Svorka = string.IsNullOrEmpty(k.Svorka) ? "SZ" : k.Svorka, 
                        Svorka = k.Popis.StartsWith("PU") ? "SZ" : k.Popis, 
                        Delka = k.Delka,
                        Popis = k.Popis
                    };
                    SeznamKabelu.Add(kopieKabelu);
                }
                //prázdný radek je pokud existují kabely
                if (z.SeznamKabelu.Count > 0)
                    SeznamKabelu.Add(new());
            }

            string directory = Path.GetDirectoryName(Cesta)!;
            string targetBase = Path.Combine(directory, "Výstup");
            if (!Directory.Exists(targetBase)) { Directory.CreateDirectory(targetBase); }
            targetBase = Path.Combine(targetBase, "Elektro.Kabely");

            Console.WriteLine($"Generování seznamu kabelů do {targetBase}.*");

            string[] sloupceKabelu = [
                nameof(Trasa.Oznaceni),
                //nameof(Trasa.KabelAll),
                nameof(Trasa.Kabel),
                nameof(Trasa.KabelVelikost),
                nameof(Trasa.Delka),
                nameof(Trasa.RozvadecAll),
                nameof(Trasa.OdkudSvorka),
                nameof(Trasa.Druh),
                nameof(Trasa.Svorka),
                //nameof(Trasa.Tag),
                //nameof(Trasa.Rozvadec),
                //nameof(Trasa.RozvadecCislo),
                //nameof(Trasa.Kabel),
                //nameof(Trasa.PocetZil),
                //nameof(Trasa.Prurezmm2),
                //nameof(Trasa.Druh),
                nameof(Trasa.Popis)
                
            ];

            SeznamKabelu.SaveToCsv(targetBase + ".csv", sloupceKabelu);
            SeznamKabelu.SaveXML(targetBase + ".xml", sloupceKabelu);
            SeznamKabelu.SaveHtmlStyle(targetBase + ".html", "Seznam vnějších spojů", sloupceKabelu);
            SeznamKabelu.SavePdfGen(targetBase + ".pdf", "Seznam vnějších spojů", sloupceKabelu);
            SeznamKabelu.SaveDocxGen(targetBase + ".docx", "Seznam vnějších spojů", sloupceKabelu);
            SeznamKabelu.SaveXlsxGen(targetBase + ".xlsx", "Seznam vnějších spojů", sloupceKabelu);
        }

        /// <summary>
        /// Zajistí, že existuje soubor Elektro.Data.json.
        /// Pokud neexistuje, nabídne uživateli dialog pro jeho výběr nebo možnost zkopírovat jej ze souboru Strojni.
        /// </summary>
        /// <returns>Cesta k souboru, nebo null, pokud se soubor nepodařilo zajistit.</returns>
        private static string? ZajistitSouborElektroJson()
        {
            var Cesta = Informace.Instance.SouborElektroJson;
            Console.WriteLine("Cesta : " + Cesta);
            if (!File.Exists(Cesta))
            {
                Cesta = Soubory.ShowOpenFileDialog("Json soubor (*.json)|*.json", Informace.Instance.BasePath);
                if (string.IsNullOrEmpty(Cesta) || !File.Exists(Cesta))
                {
                    Console.WriteLine("Výběr souboru Elektro byl stornován nebo soubor neexistuje.");

                    // Nabídneme vytvoření kopie ze souboru Strojni
                    if (MessageBox.Show("Chcete vytvořit kopii souboru ze souboru Strojni?", "Info", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        string? cestaStroje = ZajistitSouborStrojeJson();
                        if (cestaStroje == null)
                        {
                            Console.WriteLine("Nelze vytvořit kopii, protože chybí zdrojový soubor Strojni.");
                            return null;
                        }

                        Informace.Instance.SouborElektroJson = Path.Combine(Informace.Instance.BasePath, "Elektro.Data.json");
                        Informace.Instance.Ulozit();
                        Cesta = Informace.Instance.SouborElektroJson;

                        try
                        {
                            File.Copy(cestaStroje, Cesta, overwrite: true);
                            if (File.Exists(Cesta))
                            {
                                Console.WriteLine($"Soubor {Cesta} - zkopírován ze {cestaStroje}.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Chyba při kopírování souboru: {ex.Message}");
                            return null;
                        }
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
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
        private static string? ZajistitSouborStrojeJson()
        {
            var Cesta = Informace.Instance.SouborStrojeJson;
            if (!File.Exists(Cesta))
            {
                Cesta = Soubory.ShowOpenFileDialog("Json soubor (*.json)|*.json", Informace.Instance.BasePath);
                if (string.IsNullOrEmpty(Cesta) || !File.Exists(Cesta))
                {
                    Console.WriteLine("Výběr souboru strojů byl stornován nebo soubor neexistuje.");
                    return null;
                }
                else
                {
                    Informace.Instance.SouborStrojeJson = Cesta;
                    Informace.Instance.Ulozit();
                }
            }
            return Cesta;
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            //var Vývody = Path.Combine(Cesty.Elektro, "Vývody.csv");
            //var Data = Soubory.LoadFromCsv<Zarizeni>(Vývody);

            var Vývody = Cesty.VyvodyTopeniJson;
            if (!File.Exists(Vývody))
            {
                Console.WriteLine("Soubor nebyl nalezen " + Vývody);
                Console.WriteLine("Soubor bude vytvořen!");
                //var prazdny = new List<Zarizeni>();
                Soubory.SaveJson(new List<Zarizeni>(), Vývody);
                Console.WriteLine("Znovu klikni na tlačítko. Soubor byl vytvořen!");
                return;
            }
            var Data = Soubory.LoadJsonList<Zarizeni>(Vývody);
            Console.WriteLine($"Soubor {Vývody} - načten.");

            //var DataBind = new BindingList<Zarizeni>(Data);
            var table = new Table(Data);
            //SkrytSloupce(table.dataGridView1);
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

        private void DeleteNaKWToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Cesta = Informace.Instance.SouborElektroJson;
            Console.WriteLine(Cesta);
            var Data = Soubory.LoadJsonList<Zarizeni>(Cesta);
            if (Data.Count < 1)
            {
                Console.WriteLine("Soubor je prázdný " + Cesta);
            }
            List<Zarizeni> Pole = [];
            foreach (var item in Data)
            {
                if (item.Prikon == "-")
                    continue;
                if (item.Prikon == "—")
                    continue;
                
                //pokud je číslo
                if (double.TryParse(item.Prikon, out _))
                {
                    //Data.Remove(item); // smažeme ze skutečného seznamu
                    Pole.Add(item);
                }
            }
            Pole.SaveJson(Cesta);
        }

        /// <summary>
        /// Obsluha položky menu pro sloučený export všech tří seznamů do všech 6 formátů.
        /// </summary>
        /// <summary>
        /// Obsluha položky menu pro sloučený export všech tří seznamů do všech 6 formátů.
        /// </summary>
        private void SloucenySeznamToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string? cestaElektro = ZajistitSouborElektroJson();
            if (cestaElektro == null) return;

            // Načtení tří seznamů zařízení
            List<Zarizeni> hlavniElektro = Soubory.LoadJsonList<Zarizeni>(cestaElektro);
            List<Zarizeni> ostatniVyvody = Soubory.LoadJsonList<Zarizeni>(Cesty.VyvodyOstatniJson);
            List<Zarizeni> topeni = Soubory.LoadJsonList<Zarizeni>(Cesty.VyvodyTopeniJson);

            // Zabalení do sekcí pro export
            List<ExportSection<Zarizeni>> sections = [
                new ExportSection<Zarizeni> { Title = "Stroje a zařízení", Data = hlavniElektro },
                new ExportSection<Zarizeni> { Title = "Vlastní vývody mimo stroje", Data = ostatniVyvody },
                new ExportSection<Zarizeni> { Title = "Topení", Data = topeni }
            ];

            // Příprava cílového adresáře a názvu souboru
            string? adresar = Path.GetDirectoryName(cestaElektro);
            if (adresar == null) return;

            string targetBase = Path.Combine(adresar, "Výstup", "Elektro.SloucenySeznam");
            string targetDir = Path.GetDirectoryName(targetBase)!;
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            string[] sloupceZarizeni = [
                nameof(Zarizeni.Tag),
                nameof(Zarizeni.Popis),
                nameof(Zarizeni.Prikon),
                nameof(Zarizeni.Napeti),
                nameof(Zarizeni.Menic),
                nameof(Zarizeni.BalenaJednotka),
                nameof(Zarizeni.Pozice)
            ];

            Console.WriteLine($"Generování sloučeného exportu do {targetBase}.*");

            // Spuštění exportů pro všech 6 formátů
            SloucenyExporter.SaveXlsxSections(targetBase + ".xlsx", sections, "Seznam zařízení", sloupceZarizeni);
            SloucenyExporter.SaveCsvSections(targetBase + ".csv", sections, sloupceZarizeni);
            SloucenyExporter.SaveXmlSections(targetBase + ".xml", sections, sloupceZarizeni);
            SloucenyExporter.SaveHtmlSections(targetBase + ".html", sections, "Seznam zařízení", sloupceZarizeni);
            SloucenyExporter.SavePdfSections(targetBase + ".pdf", sections, "Seznam zařízení", sloupceZarizeni);
            SloucenyExporter.SaveDocxSections(targetBase + ".docx", sections, "Seznam zařízení", sloupceZarizeni);
            Console.WriteLine("Sloučený export zařízení dokončen ve všech 6 formátech!");
        }

        /// <summary>
        /// Obsluha položky menu pro sloučený export všech kabelů ze všech tří seznamů do všech 6 formátů.
        /// </summary>
        private void SloucenySeznamKabeluToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string? cestaElektro = ZajistitSouborElektroJson();
            if (cestaElektro == null) return;

            // Načtení tří seznamů zařízení
            List<Zarizeni> hlavniElektro = Soubory.LoadJsonList<Zarizeni>(cestaElektro);
            List<Zarizeni> ostatniVyvody = Soubory.LoadJsonList<Zarizeni>(Cesty.VyvodyOstatniJson);
            List<Zarizeni> topeni = Soubory.LoadJsonList<Zarizeni>(Cesty.VyvodyTopeniJson);

            // Příprava cílového adresáře
            string? adresar = Path.GetDirectoryName(cestaElektro);
            if (adresar == null) return;

            string targetKabelyBase = Path.Combine(adresar, "Výstup", "Elektro.SloucenySeznamKabelu");
            string targetDir = Path.GetDirectoryName(targetKabelyBase)!;
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            // 1. Získání tras pro jednotlivé sekce
            var trasyHlavni = ZiskejTrasyProZarizeni(hlavniElektro);
            var trasyOstatni = ZiskejTrasyProZarizeni(ostatniVyvody);
            var trasyTopeni = ZiskejTrasyProZarizeni(topeni);

            var sectionsKabely = new List<ExportSection<Trasa>> {
                new() { Title = "Kabely pro stroje a zařízení", Data = trasyHlavni },
                new() { Title = "Kabely pro vzduchotechniku", Data = trasyOstatni },
                new() { Title = "Kabely pro otopy potrůbí", Data = trasyTopeni }
            };

            string[] sloupceKabelu = [
                nameof(Trasa.Oznaceni),
                nameof(Trasa.Kabel),
                nameof(Trasa.KabelVelikost),
                nameof(Trasa.Delka),
                nameof(Trasa.RozvadecAll),
                nameof(Trasa.OdkudSvorka),
                nameof(Trasa.Druh),
                nameof(Trasa.Svorka),
                nameof(Trasa.Popis)
            ];

            Console.WriteLine($"Generování sloučeného exportu kabelů do {targetKabelyBase}.*");

            // 2. Sloučený export kabelů rozdělených do sekcí (6 formátů)
            SloucenyExporter.SaveXlsxSections(targetKabelyBase + ".xlsx", sectionsKabely, "Sloučený seznam kabelů", sloupceKabelu);
            SloucenyExporter.SaveCsvSections(targetKabelyBase + ".csv", sectionsKabely, sloupceKabelu);
            SloucenyExporter.SaveXmlSections(targetKabelyBase + ".xml", sectionsKabely, sloupceKabelu);
            SloucenyExporter.SaveHtmlSections(targetKabelyBase + ".html", sectionsKabely, "Sloučený seznam kabelů", sloupceKabelu);
            SloucenyExporter.SavePdfSections(targetKabelyBase + ".pdf", sectionsKabely, "Sloučený seznam kabelů", sloupceKabelu);
            SloucenyExporter.SaveDocxSections(targetKabelyBase + ".docx", sectionsKabely, "Sloučený seznam kabelů", sloupceKabelu);

            //KabelExcel(targetKabelyBase, cestaElektro, trasyHlavni);
            // 3. Doplnění speciálních záložek se vzorci a součty do Excelu
            //var vsechnaZarizeniProKabely = new List<Zarizeni>();
            //vsechnaZarizeniProKabely.AddRange(PripravZarizeniProKabely(hlavniElektro));
            //vsechnaZarizeniProKabely.AddRange(PripravZarizeniProKabely(ostatniVyvody));
            //vsechnaZarizeniProKabely.AddRange(PripravZarizeniProKabely(topeni));

            //var excelAppKabely = new ExcelApp(targetKabelyBase + ".xlsx");
            //List<List<string>> poleKabely = LigthChem.SeznamKabelů(vsechnaZarizeniProKabely, excelAppKabely, "Kabely Vše");
            //Pridat.Soucet(excelAppKabely, poleKabely, "Součet Kabely Vše");
            //excelAppKabely.ExcelQuit(targetKabelyBase + ".xlsx");

            Console.WriteLine("Sloučený export kabelů dokončen ve všech 6 formátech!");
        }

        public void KabelExcel(string targetKabelyBase, string cestaElektro, List<Trasa> trasa) { 
        CoverData cover = new CoverData
            {
                Zakaznik = "LUČEBNÍ ZÁVODY DRASLOVKA A.S. KOLÍN",
                Projekt = "W.005685.0200",
                Nazev = "SEMIPROVOZ IVCHS / DOKUMENTACE PRO POVOLENÍ STAVBY",
                DokumentNazev = "SOUPIS SPOTŘEBIČŮ ELEKTRO",
                Technologie = "TECHNOLOGICKÁ ELEKTROINSTALACE",
                CistyDokumentTyp = "TP-N-",
                CistyDokumentCislo = "9446",
                Revize = "A", 
                RevizeSeznam = [
                    new RevizeInfo
                    {
                        Rev = "0",
                        Date = "30.06.2026",
                        Description = "K připomínkám",
                        Stat = "PRL",
                        Zpacoval = "Tucauer",
                        Kontroloval = "Kašpar",
                        Schvalil = "Csato"
                    },
                    //new RevizeInfo
                    //{
                    //    Rev = "A",
                    //    Date = "28.06.2026",
                    //    Description = "Zapracování připomínek, finální verze",
                    //    Stat = "FIN",
                    //    Zpacoval = "Tucauer",
                    //    Kontroloval = "Kašpar",
                    //    Schvalil = "Csato"
                    //}
                ]
            };
            List<Spotrebic> spotrebice = [];
            int poradi = 1;
            foreach(var item in trasa) {
                
                Spotrebic spotrebic = new() { 
                    Polozka = poradi++.ToString(),
                    Rev = "A",
                    Pu = "PU1",
                    Umisteni = "SUŠÁRNA",
                    TechnolOznaceni = "S401.1",
                    Zarizeni = "HC SKID HEATING AND COOLING",
                    TypVelikost = "Dle dodavatele",
                    Ks = "1",
                    Pid = "PID02",
                    Rozvadec = "RM1N",
                    Napeti = "400",
                    InstalovanyPi = 2.5,
                    VypoctovyPi = 2.25, // Zadáno explicitně
                    Ivchs = "VCHS",
                    StartMotoru = "In = 4,3A",
                    Poznamka = "Společný rozváděč"
                    };
                spotrebice.Add(spotrebic);
            }

            List<Zarizeni> hlavniElektro = Soubory.LoadJsonList<Zarizeni>(cestaElektro);

            string sablonaCesta = "g:\\Můj disk\\Projekty\\SemiProvoz\\Podklady\\SeznamMichal\\Sablona.xlsx";
            new ExcelGenerator().Generuj(sablonaCesta, targetKabelyBase + "123" + ".xlsx", cover, spotrebice);
        }

        private static List<Trasa> ZiskejTrasyProZarizeni(List<Zarizeni> data)
        {
            if (data == null || data.Count == 0) return [];

            // Klonování a příprava kabelů
            var kopie = data.Select(x => Zarizeni.Clone(x)).ToList();
            var prazdne = kopie.Where(x => x.Kabel == null).ToList();
            prazdne.AddKabelCyky(1.8);
            var spolecne = kopie.Where(x => x.Kabel != null).Concat(prazdne).ToList();

            var kabelyTrida = KabelList.KabelyTrida(spolecne);
            kabelyTrida = [.. kabelyTrida.OrderBy(x => x.Hlavni.Rozvadec + x.Hlavni.RozvadecCislo)];

            var change = new List<Trasa>();
            foreach (var kabel in kabelyTrida)
            {
                if (kabel.Hlavni != null)
                    change.Add(kabel.Hlavni);
                if (kabel.PTC != null)
                    change.Add(kabel.PTC);
                if (kabel.Ovladani != null)
                    change.Add(kabel.Ovladani);
            }
            return change;
        }

        private static List<Zarizeni> PripravZarizeniProKabely(List<Zarizeni> data)
        {
            if (data == null || data.Count == 0) return [];

            var kopie = data.Select(x => Zarizeni.Clone(x)).ToList();
            var prazdne = kopie.Where(x => x.Kabel == null).ToList();
            prazdne.AddKabelCyky(1.8);
            return [.. kopie.Where(x => x.Kabel != null), .. prazdne];
        }

        private void ExcelSeznamZarizeniToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string? cestaElektro = ZajistitSouborElektroJson();
            if (cestaElektro == null) return;

            try
            {
                var zarizeniList = Soubory.LoadJsonList<Zarizeni>(cestaElektro);
                if (zarizeniList.Count == 0)
                {
                    Console.WriteLine("Žádná data k exportu.");
                    return;
                }

                CoverData cover = PripravCoverData("SOUPIS SPOTŘEBIČŮ ELEKTRO");
                List<Spotrebic> spotrebice = [];
                int poradi = 1;
                foreach (var z in zarizeniList)
                {
                    Spotrebic s = new()
                    {
                        Polozka = poradi++.ToString(),
                        Rev = "0",
                        Pu = z.Etapa,
                        Umisteni = z.Patro,
                        TechnolOznaceni = z.Tag,
                        Zarizeni = z.Popis,
                        TypVelikost = z.Typ,
                        Ks = z.Pocet.ToString(),
                        Pid = z.Pid,
                        Rozvadec = z.RozvadecOznačení,
                        Napeti = z.Napeti,
                        InstalovanyPi = double.TryParse(z.Prikon, out double p) ? p : null,
                        VypoctovyPi = null,
                        Ivchs = z.BalenaJednotka,
                        StartMotoru = z.Proud,
                        Poznamka = z.Poznamka
                    };
                    spotrebice.Add(s);
                }

                string sablonaCesta = Path.Combine(Cesty.AdresarSpusteni, "Sablony", "Sablona.xlsx");
                string targetDir = Path.Combine(Path.GetDirectoryName(cestaElektro)!, "Výstup");
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }
                string vystupCesta = Path.Combine(targetDir, "Elektro.SeznamZarizeni.xlsx");

                Console.WriteLine($"Generování seznamu zařízení do Excelu: {vystupCesta}");
                new ExcelGenerator().Generuj(sablonaCesta, vystupCesta, cover, spotrebice);
                Console.WriteLine("Generování seznamu zařízení dokončeno!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chyba při generování Excelu zařízení: {ex.Message}");
            }
        }

        private void ExcelSeznamKabeluToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string? cestaElektro = ZajistitSouborElektroJson();
            if (cestaElektro == null) return;

            try
            {
                var zarizeniList = Soubory.LoadJsonList<Zarizeni>(cestaElektro);
                var trasy = ZiskejTrasyProZarizeni(zarizeniList);
                if (trasy.Count == 0)
                {
                    Console.WriteLine("Žádná kabelová data k exportu.");
                    return;
                }

                CoverData cover = PripravCoverData("SEZNAM KABELŮ");
                List<KabelPolozka> kabely = [];
                int skupinaCislo = 1;
                bool novyVztah = true;

                foreach (var t in trasy)
                {
                    if (string.IsNullOrEmpty(t.Oznaceni) && string.IsNullOrEmpty(t.Kabel))
                    {
                        kabely.Add(new KabelPolozka());
                        novyVztah = true;
                        continue;
                    }

                    KabelPolozka kp = new()
                    {
                        Polozka = novyVztah ? skupinaCislo++.ToString() : "",
                        OznaceniKabeluPuvodni = "",
                        CisloKabelu = t.Oznaceni,
                        KabelTyp = t.Kabel,
                        Prurez = t.KabelVelikost,
                        Delka = double.TryParse(t.Delka, out double d) ? d : null,
                        ZeZarizeni = t.RozvadecAll,
                        UkonceniZe = t.OdkudSvorka,
                        DoZarizeni = t.Tag,
                        UkonceniDo = t.Svorka,
                        Poznamka = t.Popis
                    };

                    kabely.Add(kp);
                    novyVztah = false;
                }

                string sablonaCesta = Path.Combine(Cesty.AdresarSpusteni, "Sablony", "Sablona1.xlsx");
                string targetDir = Path.Combine(Path.GetDirectoryName(cestaElektro)!, "Výstup");
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }
                string vystupCesta = Path.Combine(targetDir, "Elektro.SeznamKabelu.xlsx");

                Console.WriteLine($"Generování seznamu kabelů do Excelu: {vystupCesta}");
                new ExcelGenerator().GenerujKabel(sablonaCesta, vystupCesta, cover, kabely);
                Console.WriteLine("Generování seznamu kabelů dokončeno!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chyba při generování Excelu kabelů: {ex.Message}");
            }
        }

        private CoverData PripravCoverData(string dokumentNazev)
        {
            return new CoverData
            {
                Zakaznik = "LUČEBNÍ ZÁVODY DRASLOVKA A.S. KOLÍN",
                Projekt = string.IsNullOrEmpty(Informace.Instance.Projekt) ? "W.005685.0200" : Informace.Instance.Projekt,
                Nazev = string.IsNullOrEmpty(Informace.Instance.Název) ? "SEMIPROVOZ IVCHS / DOKUMENTACE PRO POVOLENÍ STAVBY" : Informace.Instance.Název,
                DokumentNazev = dokumentNazev,
                Technologie = "TECHNOLOGICKÁ ELEKTROINSTALACE",
                CistyDokumentTyp = "TP-N-",
                CistyDokumentCislo = "9446",
                Revize = "A",
                RevizeSeznam = [
                    new RevizeInfo
                    {
                        Rev = "0",
                        Date = "30.06.2026",
                        Description = "K připomínkám",
                        Stat = "PRL",
                        Zpacoval = "Tucauer",
                        Kontroloval = "Kašpar",
                        Schvalil = "Csato"
                    }
                ]
            };
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
