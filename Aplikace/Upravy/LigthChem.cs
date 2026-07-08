using Aplikace.Sdilene;
using Aplikace.Seznam;
using Knihovna;
using Knihovna.Excel;
using Knihovna.Sdilene;
using Knihovna.Shared.Tridy;
using Knihovna.Tridy;
using System.Globalization;
using System.Reflection;
using static Knihovna.Tridy.Zarizeni;

namespace Aplikace.Upravy
{
    public static class LigthChem
    {
        /// <summary>Vytvořit ze Seznamu strojů json a Csv pro další doplnění</summary>
        public static void StrojniToJsonCsv()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr Seznamu strojů ...");
            Console.ResetColor();
            var Cesty = Informace.Instance.SouborStrojeXls;
            if (!File.Exists(Cesty))
            {
                Cesty = Soubory.ShowOpenFileDialog("Excel soubory (*.xls;*.xlsx)|*.xls;*.xlsx", Informace.Instance.BasePath);
                if (string.IsNullOrEmpty(Cesty) || !File.Exists(Cesty))
                {
                    Console.WriteLine("Výběr souboru byl stornován nebo soubor neexistuje.");
                    return;
                }
                else
                {
                    Informace.Instance.SouborStrojeXls = Cesty;
                    Informace.Instance.Ulozit();
                }

            }

            var Json = Path.ChangeExtension(Cesty, ".json");

            //AI převod stroju od strojařů do trídy
            var Stara = ExcelLoad.DataExcelInteractive(Cesty, "Seznam", 5);

            Stara.SaveJsonList(Json);

            //Jen lepší přehled dat.
            Stara.SaveToCsv(Path.ChangeExtension(Cesty, ".csv"));
        }

        public static List<Zarizeni> DwgToJson(string cesta1)
        {
            var Stara = ExcelLoad.DwgDataExcel(cesta1, "Summary", 3);

            foreach (var Data in Stara)
            {
                switch (Data.Patro)
                {
                    case "1":
                        Data.Vykres = "Xref.EM.1NP.dwg";
                        break;
                    case "2":
                        Data.Vykres = "Xref.EM.2NP.dwg";
                        break;
                    case "3":
                        Data.Vykres = "Xref.EM.3NP.dwg";
                        break;
                    case "4":
                        Data.Vykres = "Xref.EM.4NP.dwg";
                        break;
                    case "5":
                        Data.Vykres = "Xref.EM.5NP.dwg";
                        break;

                    default:
                        break;
                }
                Data.Rozvadec = "RM";
            }
            Console.WriteLine($"Načteno {Stara.Count} záznamů z {cesta1}");
            return Stara;
        }
        /// <summary>Převod extrahovaných dat z Dwg do Xls s následným převodem do Json</summary>
        public static void DwgXlsToJsonCsv()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr Dwg extrahovaných dat (např. UpravaZnovu.006.xlsm)...");
            Console.ResetColor();
            string? cesta1 = Soubory.ShowOpenFileDialog("Excel / Makro soubory (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm");
            if (string.IsNullOrEmpty(cesta1) || !File.Exists(cesta1))
            {
                Console.WriteLine("Výběr souboru byl stornován.");
                return;
            }
            var Stara = DwgToJson(cesta1);

            //Převod->json,csv 
            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Stara.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));
        }

        public static void DoplněníDat()
        {
            string cestaData = Informace.Instance.SouborElektroJson;
            if (!File.Exists(cestaData))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Otevírám dialog pro výběr stávajícího datového souboru JSON...");
                Console.ResetColor();
                cestaData = Soubory.ShowOpenFileDialog("JSON soubory (*.json)|*.json") ?? "";
                if (string.IsNullOrEmpty(cestaData)) return;
            }
            var Data = Soubory.LoadJsonList<Zarizeni>(cestaData);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr extrahovaných dat stroje (např. UpravaZnovu.006.xlsm)...");
            Console.ResetColor();
            string? cesta1 = Soubory.ShowOpenFileDialog("Excel / Makro soubory (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm");
            if (string.IsNullOrEmpty(cesta1) || !File.Exists(cesta1)) return;

            var Stroj = DwgToJson(cesta1);

            var Nove = new List<Zarizeni>();
            var Zmeny = new List<Zarizeni>();
            foreach (var item in Stroj)
            {
                var Jeden = Data.FirstOrDefault(x => x.Popis == item.Popis);
                if (Jeden == null)
                {
                    //zeznam nebyl nelezen pravděpodobně chybí
                    //záznam bude ze strojů doplněn
                    item.Nic = "Nove";
                    Nove.Add(item);
                    Zmeny.Add(item);
                }
                else
                {
                    //zaznam existuje - bude přídán již existující záznam.
                    Nove.Add(Jeden);
                }
            }
            //testovací verze
            string testCsv = Path.Combine(Path.GetDirectoryName(cesta1) ?? Informace.Instance.BasePath, "Test.csv");
            Nove.SaveToCsv(testCsv);
            //Verze přepsání původního Jsonu
        }

        /// <summary>Vytvoření excelu dle ElektroRozvaděč.Json</summary>

        private static void Preklad(List<Zarizeni> Stara)
        {
            Stara.Where(x => x.Druh == Druhy.Otop).ToList()
                .ForEach(x => x.Druh = Druhy.Otop);

            Stara.Where(x => x.Druh == Druhy.Rozvadeč).ToList()
                        .ForEach(x => x.Druh = Druhy.Rozvadeč);

            //Stara.Where(x => x.Typ.ToUpper() == "PŘÍVOD").ToList()
            Stara.Where(x => x.Typ.Equals("PŘÍVOD", StringComparison.CurrentCultureIgnoreCase)).ToList()
                        .ForEach(x => x.Typ = "Supply");

            //Stara.Where(x => x.Typ == "Spojka").ToList()
            Stara.Where(x => x.Typ.Equals("Spojka", StringComparison.CurrentCultureIgnoreCase)).ToList()
                        .ForEach(x => x.Typ = "Coupler");

            //Stara.Where(x => x.Typ == "ČERPADLO").ToList()
            Stara.Where(x => x.Typ.Equals("ČERPADLO", StringComparison.CurrentCultureIgnoreCase)).ToList()
                        .ForEach(x => x.Typ = "Pump");

            Stara.Where(x => x.Typ == "OKLEP").ToList()
                .ForEach(x => x.Typ = "Vibration");

            Stara.Where(x => x.Typ == "KOČKA").ToList()
                .ForEach(x => x.Typ = "Elevator");

            Stara.Where(x => x.Typ == "VÝVĚVA").ToList()
                .ForEach(x => x.Typ = "Vacuum");

            Stara.Where(x => x.Typ == "PODAVAČ").ToList()
                .ForEach(x => x.Typ = "Rotary");

            Stara.Where(x => x.Typ == "MÍCHADLO").ToList()
                .ForEach(x => x.Typ = "Mixer");
        }

        //public static List<List<string>> SeznamKabelů(List<Zarizeni> Stara, ExcelApp ExcelApp, string SheatName)
        //{

        //    //2.Excel založka seznam kabelů
        //    //Vytvoření pole kabelů pro zápis do Excelu

        //    {
        //    }

        //    //pole kabelů zapsat do Excel tabulky 

        //    //Nová záložka nebo nastav existující

        //    //Doplnení nadpisu a ramecku

        //    //Do Excel vyplní od radku 3 data z PoleData mělo by se jednat o seznam kabelů
        //}



        public static void AddProud()
        {
            var Stara = Soubory.LoadJsonList<Zarizeni>(Informace.Instance.SouborElektroJson);

            //Možná proud asi jen tam kde není.
            var Add = Stara.AddProud();
            Add.ToList().SaveJsonList(Informace.Instance.SouborElektroJson);
        }

        /// <summary>Seznam vývodů pro doplnění </summary>
        public static void AddVyvody()
        {
            ////Aktualní seznam vývodů

            ////Seznam vývodů pro doplnění
            //{
            //    //Pokud neexistuje tak vytvoř
            //}

            ////Načtení vývodů pro doplnnění 
            //{
            //    //přidat identifikátor
            //    { 
            //    }
            //    //
            //}

            ////Mělo by přídat proud

            ////pokud nexistuje Apid tak přepsat Json
            //{ 
            //}

            //{
            //    //hledej existeci podle Apid

            //    //Existuje tak ho smaž

            //    //znovu přidat
            //}

            ////uložit upravená data do Json

            ////Target.SaveToCsv(Path.ChangeExtension(cesta, ".csv"));
            ////Target.SaveJsonList(Path.ChangeExtension(cesta, ".txt"));
        }




        public static void DoplneniCsvToJson()
        {
            //Soubor kam bude doplněno
            string cestaData = Path.Combine(Informace.Instance.SouborElektroJson);
            var Target = Soubory.LoadJsonList<Zarizeni>(cestaData);

            //Data pro doplnění
            string cesta = Path.Combine(Cesty.ElektroDataCsv);
            if (!File.Exists(cesta)) return;
            var Source = Soubory.LoadFromCsv<Zarizeni>(cesta);

            //Doplnění dat do Json.
            Prevod.UpdateCsvToJson(Source, Target);

            //uložit doplnění informace do Json
            Target.SaveJsonList(cestaData);
        }

        public static void VyvoritFMKM()
        {
            VyvoritFM();
            VyvoritKM();
        }

        /// <summary> Převod seznamu frekvenčních měničů na Json </summary>
        public static void VyvoritFM()
        {
            string basePath = Path.Combine(Informace.Instance.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv");

            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary> Převod seznamu stykačů na Json </summary>
        public static void VyvoritKM()
        {
            string basePath = Path.Combine(Informace.Instance.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv");

            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary> Převod seznamu motorů a motorů3000 na jeden Json </summary>
        public static void VyvoritMotor()
        {
            var Motor = Soubory.LoadFromCsv<Motor>(Cesty.MotorCsv);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            var Motor3000 = Soubory.LoadFromCsv<Motor>(Cesty.Motor3000Csv);
            Console.WriteLine($"Pocet motorů: {Motor3000.Count}");

            Motor.AddRange(Motor3000);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            Motor.SaveJsonList(Cesty.Motor);
            Console.WriteLine($"Motory uloženy jako Json");
        }

        public static void Rozvadec()
        {
            var Data = Soubory.LoadJsonList<Zarizeni>(Informace.Instance.SouborElektroJson);
            var Data2 = Soubory.LoadJsonList<Zarizeni>(Cesty.VyvodyStavbaJson);

            Data = [.. Data, .. Data2];

            //rozdělení podle rozvaděče
            var skupiny = Data.GroupBy(x => x.RozvadecOznačení);
            int PocetRozvadecu = skupiny.Count();
            Console.WriteLine($"\nPočet rozvaděčů : {PocetRozvadecu}");

            foreach (var skupina in skupiny)
            {
                //prvni položka skupiny
                var Jedna = skupina.ElementAtOrDefault(1) ?? new Zarizeni();

                // Převod stringu na enum
                StringToEnum(skupina);

                //Srovnání podle enumu
                var Pole = skupina.OrderByDescending(x => double.TryParse(x.Prikon, NumberStyles.Any, CultureInfo.InvariantCulture, out double result) ? result : 0.0)
                    .OrderBy(x => x.DruhEnum).ToList();

                //Součet příkonů
                var SumaPrikon = Pole.Where(x => x.DruhEnum != Druhy.Přívod && x.DruhEnum != Druhy.Spojka)
                    .Sum(x => double.TryParse(x.Prikon, NumberStyles.Any, CultureInfo.InvariantCulture, out double result) ? result : 0.0);
                Console.WriteLine($" ");
                Console.WriteLine($"\nRozvaděč: {skupina.Key}");
                foreach (var item in Pole)
                {
                    if (item.DruhEnum == Druhy.Přívod)
                        Console.WriteLine($"Tag: {item.Tag.Replace("\n", " "),-12}, Druh: {item.Druh,-12}, Popis: {item.Popis,-35}, SumaPříkon: {SumaPrikon,-15:F2}");
                    else
                        Console.WriteLine($"Tag: {item.Tag.Replace("\n", " "),-12}, Druh: {item.Druh,-12}, Popis: {item.Popis,-35}, Příkon: {item.Prikon,-15}");
                }

            }
        }
        public static void SpojitSeznamy()
        {
            var Data = Soubory.LoadJsonList<Zarizeni>(Informace.Instance.SouborElektroJson);
            Data = [.. Data.Where(x => x.Etapa == "FAZE 1")];

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Otevírám dialog pro výběr JSON souboru FÁZE 2 k připojení (např. UpravaZnovu.006.json)...");
            Console.ResetColor();
            string? cesta1 = Soubory.ShowOpenFileDialog("JSON soubory (*.json)|*.json");
            if (string.IsNullOrEmpty(cesta1) || !File.Exists(cesta1))
            {
                Console.WriteLine("Výběr souboru byl stornován.");
                return;
            }
            var Data2 = Soubory.LoadJsonList<Zarizeni>(cesta1);
            Data2 = [.. Data2.Where(x => x.Etapa == "FAZE 2")];

            Data = [.. Data, .. Data2];
            Data.SaveJsonList(Informace.Instance.SouborElektroJson);
        }
        /// <summary>Převod stringu na enum</summary>
        private static void StringToEnum(IGrouping<string, Zarizeni> skupina)
        {
            foreach (var ukol in skupina)
            {
                if (Enum.TryParse<Druhy>(ukol.Druh.ToString(), true, out var priorita))
                {
                    ukol.DruhEnum = priorita;
                }
                else
                {
                    ukol.DruhEnum = Druhy.Nic; // nebo jiná výchozí hodnota
                }
            }
        }

        internal static void Duplicity()
        {
            string cestaData = Informace.Instance.SouborElektroJson;
            var Data = Soubory.LoadJsonList<Zarizeni>(cestaData);
            Console.WriteLine($"Pocet záznamů: {Data.Count}");
            Data = [.. Data.DistinctBy(x => x.Apid)];

            //Verze přepsání původního Jsonu
            Data.SaveJsonList(cestaData);
        }

        /// <summary> Asi se jedná o náčítání seznamu výkresů z excelu pro další zpracování</summary>
        /// 17 je rádek kde to má začínat načítat
        /// V souboru pak vypsáno co se kam má načítat.
        internal static void NačtiSeznamVýkresůXls(string v)
        {
            var data = ExcelLoad.DataExcelVykres(v, "seznam dokumentace", 17);
            data.SaveJsonList(Path.ChangeExtension(v, ".json"));
        }
    }
}

