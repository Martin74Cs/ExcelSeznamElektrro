// See https://aka.ms/new-console-template for more information
using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using Aplikace.Upravy;


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
while (!konec)
{
    using var currentInfo = Informace.Create;
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
    Console.WriteLine("1. Načíst seznam výkresů z XLS (Lightchem)");
    Console.WriteLine("2. Spustit kompletní Elektro / Revizní proces");
    Console.WriteLine("3. Zpracovat místnosti (Vytvořit seznamy)");
    Console.WriteLine("4. Zpracovat Povrly (JSON -> XML/CSV)");
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("5. Zobrazit podrobnou nápovědu a strukturu aplikace");
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("6. Nastavit složku projektu");
    Console.ResetColor();
    Console.WriteLine("0. Konec");
    Console.WriteLine("------------------------------------------------------------------");
    Console.Write("Vyberte možnost [0-6]: ");

    //Nastavení cesty 
    //Data.Instance.Cesta = currentPath;

    string? volba = Console.ReadLine();
    Console.WriteLine();

    try
    {
        switch (volba)
        {
            case "1":
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Otevírám dialog pro výběr XLS/XLSX výkresového souboru...");
                Console.ResetColor();
                string? cesta = Soubory.ShowOpenFileDialog("Excel soubory (*.xls;*.xlsx)|*.xls;*.xlsx");
                if (string.IsNullOrEmpty(cesta))
                {
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
                if (!string.IsNullOrEmpty(novaCesta))
                {
                    var info = Informace.Create;
                    info.BasePath = novaCesta;
                    info.Ulozit();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Projektová složka byla úspěšně změněna na: {novaCesta}");
                    Console.ResetColor();
                }
                else
                {
                    Console.WriteLine("Změna byla zrušena.");
                }
                Console.WriteLine("\nStiskněte libovolnou klávesu...");
                Console.ReadKey();
                break;

            case "0":
                konec = true;
                //Informace.Create.Ulozit();
                Console.WriteLine("Ukončuji aplikaci. Na shledanou!");
                break;

            default:
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Neplatná volba, zkuste to znovu.");
                Console.ResetColor();
                System.Threading.Thread.Sleep(1000);
                break;
        }
    }
    catch (Exception ex)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Došlo k chybě při provádění operace: {ex.Message}");
        Console.ResetColor();
        Console.WriteLine("\nStiskněte libovolnou klávesu...");
        Console.ReadKey();
    }
}

void ZobrazitNapovedu()
{
    Console.Clear();
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("==================================================================");
    Console.WriteLine("                 NÁPOVĚDA A DOKUMENTACE APLIKACE                  ");
    Console.WriteLine("==================================================================");
    Console.ResetColor();
    Console.WriteLine("Tato aplikace slouží k automatizaci zpracování elektroseznamů,");
    Console.WriteLine("kabelových specifikací, seznamů místností a fluidních zařízení.");
    Console.WriteLine();
    Console.WriteLine("PŮVODNÍ NÁPOVĚDA A METODY APLIKACE:");
    Console.WriteLine("- Soubory.KillExcel(): Ukončí všechny běžící procesy excel.exe");
    Console.WriteLine("  (nyní již není nutné, protože aplikace používá ClosedXML,");
    Console.WriteLine("  který nepotřebuje běžící Excel na pozadí).");
    Console.WriteLine("- ElektroLoad.NovyExcel(): Založí nový čistý Excel dokument.");
    Console.WriteLine("- LigthChem.Hlavni() & LigthChem.Rozvadec():");
    Console.WriteLine("  Slouží pro specifické zpracování rozvaděčů a seznamů.");
    Console.WriteLine("- LigthChem.NačtiSeznamVýkresůXls(cesta):");
    Console.WriteLine("  Načte seznam výkresů z XLS a převede do struktury projektu.");
    Console.WriteLine();
    Console.WriteLine("STRUKTURA PROJEKTU A SLOŽKY:");
    Console.WriteLine("- Aplikace/Excel: Obsahuje 'ExcelApp.cs', což je spravovaný");
    Console.WriteLine("  wrapper nad ClosedXML nahrazující starý COM interop.");
    Console.WriteLine("- Aplikace/Sdilene: Pomocné metody pro výpočty proudů motorů");
    Console.WriteLine("  (Pridat.cs), pracování s cestami a konfigurací.");
    Console.WriteLine("- Aplikace/Seznam: Načítání a porovnávání revizí elektro prvků.");
    Console.WriteLine();
    Console.WriteLine("LOGOVÁNÍ A KÓDOVÁNÍ:");
    Console.WriteLine("- Výstupy z konzole se automaticky ukládají do souboru 'console.log'");
    Console.WriteLine("  v kódování Windows-1250 (vhodné pro Windows).");
    Console.WriteLine("==================================================================");
    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine("Stiskněte libovolnou klávesu pro návrat do hlavní nabídky...");
    Console.ResetColor();
    Console.ReadKey();
}

public class DoubleWriter : System.IO.TextWriter
{
    private readonly System.IO.TextWriter _w1;
    private readonly System.IO.TextWriter _w2;

    public DoubleWriter(System.IO.TextWriter w1, System.IO.TextWriter w2)
    {
        _w1 = w1;
        _w2 = w2;
    }

    public override System.Text.Encoding Encoding => _w1.Encoding;

    public override void Write(char value)
    {
        _w1.Write(value);
        _w2.Write(value);
    }

    public override void Write(string? value)
    {
        _w1.Write(value);
        _w2.Write(value);
    }

    public override void Flush()
    {
        _w1.Flush();
        _w2.Flush();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _w1.Dispose();
            _w2.Dispose();
        }
        base.Dispose(disposing);
    }
}

