// See https://aka.ms/new-console-template for more information
using Aplikace;
using Aplikace.Seznam;
using Aplikace.Upravy;
using Knihovna;
using Knihovna.Export;
using Knihovna.Sdilene;
using Knihovna.Tridy;

using Parametr.MAcad;

// Setup console logging to a file in Windows-1250 encoding
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
var win1250 = System.Text.Encoding.GetEncoding(1250);
var logWriter = new System.IO.StreamWriter("console.log", append: false, encoding: win1250) { AutoFlush = true };
var doubleWriter = new DoubleWriter(Console.Out, logWriter);
Console.SetOut(doubleWriter);
AppDomain.CurrentDomain.ProcessExit += (s, e) => {
    doubleWriter.Flush();
    logWriter.Dispose();
};

bool konec = false;
while(!konec) {
    var currentInfo = Informace.Instance;
    string currentPath = string.IsNullOrEmpty(currentInfo.BasePath) ? "[NENASTAVENO]" : currentInfo.BasePath;

    Console.Clear();
    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine("==================================================================");
    Console.WriteLine("        ELEKTRO SEZNAMY A EXCEL EXPORTY - HLAVNÍ NABÍDKA         ");
    Console.WriteLine("==================================================================");
    Console.ResetColor();
    Console.ForegroundColor = ConsoleColor.Gray;
    Console.WriteLine($"Aktuální složka projektu: {currentPath}");
    Console.WriteLine("------------------------------------------------------------------");
    Console.ResetColor();
    Console.WriteLine("1. Načíst seznam výkresů z XLS");
    Console.WriteLine("2. Spustit kompletní Elektro / Revizní proces");
    Console.WriteLine("3. Zpracovat místnosti (Vytvořit seznamy)");
    Console.WriteLine("4. Zpracovat Povrly (JSON -> XML/CSV)");
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("5. Zobrazit podrobnou nápovědu a strukturu aplikace");
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("6. Nastavit složku projektu");
    Console.ForegroundColor = ConsoleColor.Magenta;
    Console.WriteLine("7. Test nahrat a uložit json");
    Console.ResetColor();
    Console.WriteLine("0. Konec");
    Console.WriteLine("------------------------------------------------------------------");
    Console.Write("Vyberte možnost [0-100]: ");

    //Nastavení cesty 

    string? volba = Console.ReadLine();
    Console.WriteLine();

    try {
        switch(volba) {
            case "1":
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Otevírám dialog pro výběr XLS/XLSX výkresového souboru...");
                Console.ResetColor();
                string? cesta = Soubory.ShowOpenFileDialog("Excel soubory (*.xls;*.xlsx)|*.xls;*.xlsx");
                if(string.IsNullOrEmpty(cesta)) {
                    Console.WriteLine("Výběr souboru byl stornován.");
                    Console.WriteLine("\nStiskněte libovolnou klávesu...");
                    Console.ReadKey();
                    break;
                }
                Console.WriteLine($"Spouštím: NačtiSeznamVýkresůXls pro soubor: {cesta}");
                LigthChem.NačtiSeznamVýkresůXls(cesta);
                Console.WriteLine("\nHotovo. Stiskněte libovolnou klávesu...");
                Console.ReadKey();
                break;

            case "2":
                Console.WriteLine("Spouštím kompletní Elektro / Revizní proces...");
                ElektroLoad.Elektro();
                Console.WriteLine("\nHotovo. Stiskněte libovolnou klávesu...");
                Console.ReadKey();
                break;

            case "3":
                Console.WriteLine("Spouštím zpracování místností (Místnosti.VytvoritSeznamy)...");
                Místnosti.VytvoritSeznamy();
                Console.WriteLine("\nHotovo. Stiskněte libovolnou klávesu...");
                Console.ReadKey();
                break;

            case "4":
                Console.WriteLine("Spouštím zpracování Povrly (Povrly.Hlavni)...");
                Povrly.Hlavni();
                Console.WriteLine("\nHotovo. Stiskněte libovolnou klávesu...");
                Console.ReadKey();
                break;

            case "5":
                ZobrazitNapovedu();
                break;

            case "6":
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("--- NASTAVENÍ SLOŽKY PROJEKTU ---");
                Console.ResetColor();
                Console.WriteLine($"Současná složka: {currentPath}");
                Console.WriteLine("Otevírám dialog pro výběr složky...");
                string? novaCesta = Soubory.ShowFolderBrowserDialog("Vyberte hlavní složku projektu", currentPath);
                if(!string.IsNullOrEmpty(novaCesta)) {
                    var info = Informace.Instance;
                    info.BasePath = novaCesta;
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Projektová složka byla úspěšně změněna na: {novaCesta}");
                    Console.ResetColor();
                    Informace.Instance.BasePath = novaCesta;
                    Informace.Instance.Ulozit();
                }
                else {
                    Console.WriteLine("Změna byla zrušena.");
                }
                Console.WriteLine("\nStiskněte libovolnou klávesu...");
                Console.ReadKey();
                break;

            case "7":
                Console.WriteLine($"\nCesta : {Informace.Instance.BasePath} ");
                string CestaSoubor = Soubory.ShowOpenFileDialog("(*.json)|*.json", Informace.Instance.BasePath)??"";
                if (string.IsNullOrEmpty(CestaSoubor))  {
                    Console.WriteLine("Výběr souboru byl stornován.");
                    return;
                }
                var Json = Soubory.LoadJsonEn<SumoResult>(CestaSoubor);
                Console.WriteLine($"\nNačteno. {Json.Count} záznamů");
                foreach(var item in Json) {
                    Console.WriteLine($"Jmeno zařízení : {item.Text}, {item.KksCount},{item.SumoHandle} ");
                }

                var FlatJson = FlatRow.Flatten(Json);
                FlatJson.SaveHtmlStyleFlat(Path.ChangeExtension(CestaSoubor, ".html"));
                Json.SaveXML(Path.ChangeExtension(CestaSoubor, ".xml"));

                //DocxGenerator
                FlatJson.SaveDocxGenFlat(Path.ChangeExtension(CestaSoubor, ".docx"));

                FlatJson.SavePdfGenFlat(Path.ChangeExtension(CestaSoubor, ".pdf"));

                FlatJson.SaveToCsv(Path.ChangeExtension(CestaSoubor, ".csv"));

                Console.WriteLine("\nHotovo. Stiskněte libovolnou klávesu...");
                Console.ReadKey();
                break;

            case "8":
                SumoKKs();
                break;

            case "0":
                konec = true;
                Console.WriteLine("Ukončuji aplikaci. Na shledanou!");
                break;

            default:
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Neplatná volba, zkuste to znovu.");
                Console.ResetColor();
                System.Threading.Thread.Sleep(1000);
                break;
        }
    } catch(Exception ex) {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Došlo k chybě při provádění operace: {ex.Message}");
        Console.ResetColor();
        Console.WriteLine("\nStiskněte libovolnou klávesu...");
        Console.ReadKey();
    }
}

static void SumoKKs() {
    Console.WriteLine($"\nCesta : {Informace.Instance.BasePath} ");
    string CestaSoubor = Soubory.ShowOpenFileDialog("(*.json)|*.json", Informace.Instance.BasePath) ?? "";
    if (string.IsNullOrEmpty(CestaSoubor))
    {
        Console.WriteLine("Výběr souboru byl stornován.");
        return;
    }
    var Json = Soubory.LoadJsonEn<SumoDivisionLog>(CestaSoubor);
    Console.WriteLine($"\nNačteno. {Json.Count} záznamů");


    var FlatJson = Json.ToFlatRows();
    FlatJson.SaveHtmlStyleFlat(Path.ChangeExtension(CestaSoubor, ".html"));
    Json.SaveXML(Path.ChangeExtension(CestaSoubor, ".xml"));

    //DocxGenerator
    FlatJson.SaveDocxGenFlat(Path.ChangeExtension(CestaSoubor, ".docx"));

    FlatJson.SavePdfGenFlat(Path.ChangeExtension(CestaSoubor, ".pdf"));

    FlatJson.SaveToCsv(Path.ChangeExtension(CestaSoubor, ".csv"));

    Console.WriteLine("\nHotovo. Stiskněte libovolnou klávesu...");
    Console.ReadKey();
}

static void ZobrazitNapovedu() {
    Console.Clear();
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("==================================================================");
    Console.WriteLine("                 NÁPOVĚDA A DOKUMENTACE APLIKACE                  ");
    Console.WriteLine("==================================================================");
    Console.ResetColor();
    Console.WriteLine("Tato aplikace slouží k automatizaci zpracování elektroseznamů,");
    Console.WriteLine("kabelových specifikací, seznamů místností a fluidních zařízení.");
    Console.WriteLine();
    Console.WriteLine("ZÁKLADNÍ POUŽITÍ A STRUKTURA:");
    Console.WriteLine("1. Nastavení složky projektu (Volba 6): Nastaví základní adresář");
    Console.WriteLine("   projektu (uloží se do %APPDATA%/Elektro/data.txt).");
    Console.WriteLine("   Podle tohoto adresáře se odvozují všechny cesty k souborům");
    Console.WriteLine("   a výstupům (které se ukládají do podsložky 'Výstup').");
    Console.WriteLine("2. Načtení seznamu výkresů (Volba 1): Načte seznam výkresů z XLS.");
    Console.WriteLine("3. Spustit kompletní proces (Volba 2): Spustí revizní a elektro proces.");
    Console.WriteLine("4. Zpracovat místnosti (Volba 3): Vytvoří seznamy místností.");
    Console.WriteLine("5. Zdroje dat: Aplikace načítá specifikace motorů, kabelů, stykačů");
    Console.WriteLine("   a jističů z JSON souborů. Pokud není nastavena externí složka,");
    Console.WriteLine("   použije se lokální složka 'ZdrojeDat' v adresáři aplikace.");
    Console.WriteLine();
    Console.WriteLine("OŠETŘENÍ PŘEPISOVÁNÍ SOUBORŮ:");
    Console.WriteLine("- Při jakémkoliv exportu nebo zápisu souboru na disk program ověřuje,");
    Console.WriteLine("  zda již soubor existuje. Pokud ano, v konzolové verzi se dotáže");
    Console.WriteLine("  uživatele přes textové rozhraní, ve WinForms verzi zobrazí dialog.");
    Console.WriteLine();
    Console.WriteLine("PŮVODNÍ METODY:");
    Console.WriteLine("- Soubory.KillExcel(): Ukončí běžící procesy excel.exe.");
    Console.WriteLine("- ElektroLoad.NovyExcel(): Založí čistý Excel dokument.");
    Console.WriteLine("==================================================================");
    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine("Stiskněte libovolnou klávesu pro návrat do hlavní nabídky...");
    Console.ResetColor();
    Console.ReadKey();
}

namespace Aplikace {

    public class DoubleWriter(System.IO.TextWriter w1, System.IO.TextWriter w2): System.IO.TextWriter {
        private readonly System.IO.TextWriter _w1 = w1;
        private readonly System.IO.TextWriter _w2 = w2;

        public override System.Text.Encoding Encoding => _w1.Encoding;

        public override void Write(char value) {
            _w1.Write(value);
            _w2.Write(value);
        }

        public override void Write(string? value) {
            _w1.Write(value);
            _w2.Write(value);
        }

        public override void Flush() {
            _w1.Flush();
            _w2.Flush();
        }

        protected override void Dispose(bool disposing) {
            if(disposing) {
                _w1.Dispose();
                _w2.Dispose();
            }
            base.Dispose(disposing);
        }
    }

}


