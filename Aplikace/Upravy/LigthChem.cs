using Aplikace.Excel;
using Aplikace.Sdilene;
using Aplikace.Seznam;
using Aplikace.Tridy;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using static Aplikace.Tridy.Motor;
using static Aplikace.Tridy.Zarizeni;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Exc = Microsoft.Office.Interop.Excel;

namespace Aplikace.Upravy
{
    public static class LigthChem
    {
        /// <summary>Vytvořit z Seznamu strojů json a Csv pro další doplnění</summary>
        public static void StrojniToJsonCsv()
        {
            //Převod->json,csv
            
            //string basePath = @"G:\z\W.002115_NATRON\Prac_Prof\e_EL\vykresy\Martin_PRS\2024.09.03";
            if (!Directory.Exists(Cesty.Elektro))
                Directory.CreateDirectory(Cesty.Elektro);

            string cesta1 = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            var Stara = ExcelLoad.DataExcel(cesta1, "Seznam", 8);

            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Stara.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));

            //string cestaData = Path.Combine(Cesty.ElektroDataCsv);
            //if (!File.Exists(cestaData))
            //{
            //    //vytvoření základu pro json jen pokud neexistuje.
            //    Stara.SaveJsonList(Path.ChangeExtension(cestaData, ".json"));
            //    //vytvoření csv pro doplnění
            //    Stara.SaveToCsv(cestaData);
            //}
            //else
            //{
            //    Console.WriteLine($"Vyvoření kopie {Path.GetFileName(cestaData)} přeskočeno");
            //}
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
            string cesta1 = Path.Combine(Cesty.Elektro, "Pid", @"UpravaZnovu.006.xlsm");
            var Stara = DwgToJson(cesta1);

            //Převod->json,csv 
            Stara.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            Stara.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));
        }

        public static void DoplněníDat()
        {
            string cestaData = Cesty.ElektroDataJson;
            var Data = Soubory.LoadJsonList<Zarizeni>(cestaData);

            string cesta1 = Path.Combine(Cesty.Elektro, "Pid", @"UpravaZnovu.006.xlsm");
            var Stroj = DwgToJson(cesta1);

            var Nove = new List<Zarizeni>();
            var Zmeny = new List<Zarizeni>();
            foreach (var item in Stroj)
            {
                var Jeden = Data.FirstOrDefault(x => x.Popis == item.Popis);
                //var Jeden = Data.FirstOrDefault(x => x.Etapa == "FAZE 2");
                if (Jeden == null)
                {
                    //zeznam nebyl nelezen pravděpodobně chybí
                    //záznam bude ze strojů doplněn
                    item.Nic = "Nove";
                    Nove.Add(item);
                    Zmeny.Add(item);
                }
                else {
                    //zaznam existuje - bude přídán již existující záznam.
                    Nove.Add(Jeden);
                }
            }
            //testovací verze
            Nove.SaveToCsv(Path.Combine(Cesty.Elektro, "Pid", "Test.csv"));
            //Verze přepsání původního Jsonu
            //Nove.SaveJsonList(cestaData);
        }

        /// <summary>Vytvoření excelu dle ElektroRozvaděč.Json</summary>
        public static void JsonToExcel() {
            string cestaData = Cesty.ElektroDataJson;
            var Stara = Soubory.LoadJsonList<Zarizeni>(cestaData);

            //Vývody pro doplnění
            //var cestaVývody = Path.Combine(Cesty.Elektro, "Vývody.json");
            //var Vývody = Soubory.LoadJsonList<Zarizeni>(cestaVývody);

            //Spojení ZAŘÍZENÍ A ROZVADĚČE
            //Stara = [.. Stara, .. Vývody];

            //Vývody pro doplnění
            var cestaStavba = Path.Combine(Cesty.Elektro, "Vývody.Stavba.json");
            var VývodyStavba = Soubory.LoadJsonList<Zarizeni>(cestaStavba);

            //Spojení původních s doplněnými
            Stara = [.. Stara, .. VývodyStavba];

            //Překlad do Angličtiny
            Preklad(Stara);

            Console.WriteLine($"Načteno {Stara.Count} záznamů z {cestaData}");
            Stara = [.. Stara.Where(x => x.Etapa == "FAZE 1")];
            Console.WriteLine($"Pouze záznamy FAZE 1");

            //Možná proud asi jen tam kde není.
            //Stara.AddProud();

            //Přidání typu kabelu pokud chybí.
            var prazdne = Stara.Where(x => x.Kabel == null).ToList();
            prazdne.AddKabelCyky(1.8);

            // Spojení původních neprázdných s doplněnými kabely – Concat vytvoří novou spojenou kolekci
            //Stara = Stara.Where(x => x.Kabel != null).Concat(prazdne).ToList();
            Stara = [.. Stara.Where(x => x.Kabel != null), .. prazdne];

            var StaraVSD = Stara.Where(x => x.Menic == "VSD").ToList();
            var StaraBezVSD = Stara.Where(x => x.Menic != "VSD").ToList(); // vyloučí všechny s Menic == "VSD"

            string filename = "Seznam.xlsx";
            var cesta = Path.Combine(Cesty.Elektro, filename);
            //vytvoření nebo otevření dokumentu elektro
            var ExcelApp = new ExcelApp(cesta);

            //1.Excel založka seznam kabelů
            //Nastavení nebo vytvoření záložky

            //čísla sloupců a nazvy tříd. 
            var dir = new Dictionary<int, string>() {
                {1,"Tag"},
                {2,"Popis"},
                {3,"Prikon"},
                {4,"Napeti"},
                {5,"Druh"},
                {6,"Menic"},
                {7,"Typ"},
                {8,"RozvadecOznačení"},
                {9,"Predmet"},

                //{1,"Tag"},
                //{2,"Popis"},
                //{3,"Prikon"},
                //{4,"Napeti"},
                ////{5,"Proud"},
                //{6,"Druh"},
                //{7,"Menic"},
                //{8,"Typ"},
                ////{9,"PrurezMM2"},
                ////{10,"Delka"},
                //{11,"RozvadecOznačení"},
                ////{12,"RozvadecCislo"},
                //{12,"Predmet"},

                //{100,"PID"},
                //{101,"Nic"},
                //{102,"HP"},
                //{103,"AWG"},
                //{105,"Napeti"},
                //{106,"Radek"},
            };

            ExcelApp.GetSheet("Seznam Všechno");
            //Vytvoření nadpisů
            var range = ExcelApp.Nadpisy([.. Nadpis.DataEn()]);
            //Formátování nadpisů
            ExcelApp.NadpisSet(range);
            ExcelApp.ClassToExcel(Row: 3, Stara, dir);

            ExcelApp.GetSheet("Seznam Elektro");
            //Vytvoření nadpisů
            range = ExcelApp.Nadpisy([.. Nadpis.DataEn()]);
            //Formátování nadpisů
            ExcelApp.NadpisSet(range);
            ExcelApp.ClassToExcel(Row: 3, StaraBezVSD, dir);

            //2.Excel založka seznam kabelů
            ExcelApp.GetSheet("Seznam Elektro VSD");
            range = ExcelApp.Nadpisy([.. Nadpis.DataEn()]);
            ExcelApp.NadpisSet(range);
            ExcelApp.ClassToExcel(Row: 3, StaraVSD, dir);

            //3.Excel založka seznam kabelů
            //Vyzváření seznamu kabelů podle krytérii
            List<List<string>> PoleDataBezVSD = SeznamKabelů(StaraBezVSD, ExcelApp, "Kabely");
            //4.Excel založka seznam kabelů
            Pridat.Soucet(ExcelApp, PoleDataBezVSD, "Součet Kabely");

            //5.Excel založka seznam kabelů
            //Vyzváření seznamu kabelů podle krytérii
            var PoleDataVSD = SeznamKabelů(StaraVSD, ExcelApp, "Kabely VSD");
            //6.Excel založka seznam kabelů
            Pridat.Soucet(ExcelApp, PoleDataVSD, "Seznam VSD");

            ExcelApp.ExcelQuit(cesta);

        }

        private static void Preklad(List<Zarizeni> Stara) {
            Stara.Where(x => x.Druh == "Otop").ToList()
                        .ForEach(x => x.Druh = "Heating");

            Stara.Where(x => x.Druh == "Rozvadeč").ToList()
                        .ForEach(x => x.Druh = "Distributor");

            Stara.Where(x => x.Typ.ToUpper() == "PŘÍVOD").ToList()
                        .ForEach(x => x.Typ = "Supply");

            Stara.Where(x => x.Typ == "Spojka").ToList()
                        .ForEach(x => x.Typ = "Coupler");

            Stara.Where(x => x.Typ == "ČERPADLO").ToList()
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

        private static List<List<string>> SeznamKabelů(List<Zarizeni> Stara, ExcelApp ExcelApp, string SheatName)
        {
            Console.WriteLine("Probíhá načítaní kabelů");

            //2.Excel založka seznam kabelů
            //Vytvoření pole kabelů pro zápis do Excelu
            var DataTrida = KabelList.KabelyTrida(Stara);
            DataTrida = [.. DataTrida.OrderBy(x => x.Hlavni.Rozvadec + x.Hlavni.RozvadecCislo)];

            var Change = new List<Trasa>();
            foreach (var kabel in DataTrida)
            {
                if (kabel.Hlavni != null)
                    Change.Add(kabel.Hlavni);
                if (kabel.PTC != null)
                    Change.Add(kabel.PTC);
                if (kabel.Ovladani != null)
                    Change.Add(kabel.Ovladani);
            }

            //pole kabelů zapsat do Excel tabulky 
            var PoleData = KabelList.KabelyTridaToString(Change);
            //var PoleData = KabelList.Kabely(Stara);

            //Nová záložka nebo nastav existující
            ExcelApp.GetSheet(SheatName);

            //Doplnení nadpisu a ramecku
            //ExcelApp.ExcelSaveNadpis(PoleData);
            ExcelApp.ExcelSaveNadpisEn(PoleData);

            //Do Excel vyplní od radku 3 data z PoleData mělo by se jednat o seznam kabelů
            ExcelApp.KabelyToExcel(PoleData, 3);
            return PoleData;
        }



        public static void Hlavni()
        {
            string cesta = Path.Combine(Cesty.Elektro, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
            string json = Path.ChangeExtension(cesta, ".json");
            if (!File.Exists(json))
                return;

            //var Source = Soubory.LoadFromCsv<Zarizeni>(cesta);
            //Prevod.UpdateCsvToJson(Source, Target);
        }

        public static void AddProud()
        {
            string cestaData = Cesty.ElektroDataJson;
            var Stara = Soubory.LoadJsonList<Zarizeni>(cestaData);

            //Možná proud asi jen tam kde není.
            var Add = Stara.AddProud();
            Add.ToList().SaveJsonList(cestaData);
        }

        /// <summary>Seznam vývodů pro doplnění </summary>
        public static void AddVyvody()
        {
            ////Aktualní seznam vývodů
            //string cesta = Cesty.ElektroDataJson;
            //var Target = Soubory.LoadJsonList<Zarizeni>(cesta);

            ////Seznam vývodů pro doplnění
            //var Add = new List<Zarizeni>();
            //string cesta1 = Path.Combine(Cesty.Elektro, @"Vývody.csv");
            //if (!File.Exists(cesta1))
            //{
            //    //Pokud neexistuje tak vytvoř
            //    Add.Add(new Zarizeni());
            //    Add.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            //    Add.SaveToCsv(cesta1);
            //    return;
            //}

            ////Načtení vývodů pro doplnnění 
            //Add = Soubory.LoadFromCsv<Zarizeni>(cesta1);
            //bool zmena = false;
            //for (int i = 0; i < Add.Count; i++)
            //{
            //    //přidat identifikátor
            //    if (string.IsNullOrEmpty(Add[i].Apid))
            //    { 
            //        Add[i].Apid = ExcelLoad.Apid();
            //        zmena = true;
            //    }
            //    //
            //}

            ////Mělo by přídat proud
            //Add.AddProud();

            ////pokud nexistuje Apid tak přepsat Json
            //if (zmena)
            //{ 
            //    //Prevod.SaveToCsv(Add, cesta1);
            //    //Add.SaveToCsv(cesta1);
            //    Add.SaveJsonList(Path.ChangeExtension(cesta1, ".json"));
            //}

            //int pocet = Target.Count + 1;
            //foreach (var item in Add)
            //{
            //    //hledej existeci podle Apid
            //    var Toto = Target.FirstOrDefault(x => x.Apid == item.Apid);

            //    //Existuje tak ho smaž
            //    if (Toto != null) //continue;
            //        Target.Remove(Toto);

            //    //znovu přidat
            //    item.Radek = pocet++;
            //    Target.Add(item);
            //}

            ////uložit upravená data do Json
            //Target.SaveJsonList(cesta);

            ////Target.SaveToCsv(Path.ChangeExtension(cesta, ".csv"));
            ////Target.SaveJsonList(Path.ChangeExtension(cesta, ".txt"));
        }

        //public static void AddKabely()
        //{
        //    //string cesta1 = Path.Combine(basePath, @"N92120_Seznam_stroju_zarizeni_250311_250407.xlsx");
        //    string cesta1 = Path.Combine(Cesty.ElektroDataJson);
        //    var Target = Soubory.LoadJsonList<Zarizeni>(cesta1);

        //    Target.AddKabelCyky(1.6);

        //    Target.SaveJsonList(cesta1);
        //    //Target.SaveToCsv(Path.ChangeExtension(cesta1, ".csv"));
        //}

        public static IEnumerable<Zarizeni> AddKabelCyky(this IEnumerable<Zarizeni> target, double rezerva = 1.5)
        {
            var Target = target.ToList(); // převod na List, abys mohl indexovat

            var KabelCu = Soubory.LoadJsonListEn<KabelVse>(Cesty.CuJson)
                //.Where(x => x.Name.Contains("CYKY", StringComparison.OrdinalIgnoreCase) && x.Deleni == "4")
                .Where(x => x.Name.Contains("PraFlaDur ", StringComparison.OrdinalIgnoreCase) && x.Deleni == "3")
                //.Where(x => x.Name == "PRAFlaDur" && x.Deleni == "4")
                .OrderBy(x => x.IzAE).ToList();
            //vymazaní položek ze seznam
            KabelCu.RemoveAll(x => x.SLmm2 < 2.5);

            var KabelCYKY = Soubory.LoadJsonListEn<KabelVse>(Cesty.CuJson)
                .Where(x => x.Name.Contains("CYKY", StringComparison.OrdinalIgnoreCase) && x.Deleni == "4")
                .OrderBy(x => x.IzAE).ToList();
            //vymazaní položek ze seznam
            KabelCYKY.RemoveAll(x => x.SLmm2 < 2.5);

            if (KabelCu.Count < 1) Console.WriteLine("Chyba hledání kabelů Cu");
            else Console.WriteLine($"Kabelů Cu je {KabelCu.Count}");

            //Kontrola
            //KabelCu.SaveJsonList(Path.ChangeExtension(cesta1, ".txt"));

            var KabelAL = Soubory.LoadJsonList<KabelVse>(Cesty.AlJson).OrderByDescending(x => x.IzAE ).ToList(); 
            //vymazaní položek ze seznam
            KabelAL.RemoveAll(x => x.SLmm2 < 16);
            if(KabelAL.Count < 1) Console.WriteLine("Chyba hledání kabelů Al");
            else Console.WriteLine($"Kabelů Al je {KabelAL.Count}");

            var properties = typeof(Zarizeni).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => p.CanWrite && p.Name != "Item")
                                 .ToList();

            for(int i = 0; i < Target.Count; i++)
            {
                int pocet = 1;
                bool volba = true;
                while (volba)
                {
                    foreach (var prop in properties)
                    {
                        //var value = prop.GetValue(target);
                        var proud = double.TryParse(Target[i].Proud, out var p) ? p : 1;
                        //var JedenKabel = KabelCu.FirstOrDefault(x => x.MaxProudVzduch > proud * rezerva) ?? KabelCu.FirstOrDefault(x => pocet * x.MaxProud > proud * rezerva);

                        if (proud < 150)
                        {
                            //kabel PraFlaDur
                            var JedenKabel = KabelCu.FirstOrDefault(x => x.MaxProudVzduch > proud * rezerva);
                            JedenKabel ??= KabelCu.FirstOrDefault(x => pocet * x.MaxProud > proud * rezerva);
                            if (JedenKabel == null) continue;

                            //prop.SetValue(target, value);
                            //target.GetType().GetProperty("PrurezMM2").SetValue(target, JedenKabel.SLmm2.Tostring());
                            if (pocet > 1) {
                                Target[i].PrurezMM2 = JedenKabel.SLmm2.ToString();
                                Target[i].PocetKabelu = pocet;
                            }
                            
                            else Target[i].PrurezMM2 = JedenKabel.SLmm2.ToString();
                            Target[i].Vodice = JedenKabel.Deleni;
                            Target[i].Kabel = JedenKabel;
                            volba = false;
                            break;
                        }
                        else {
                            //kabel CYKY
                            var JedenKabel = KabelCYKY.FirstOrDefault(x => x.MaxProudVzduch > proud * rezerva);
                            JedenKabel ??= KabelCYKY.FirstOrDefault(x => pocet * x.MaxProud > proud * rezerva);
                            if (JedenKabel == null) {
                                pocet++;
                            continue;
                            }

                            //prop.SetValue(target, value);
                            //target.GetType().GetProperty("PrurezMM2").SetValue(target, JedenKabel.SLmm2.Tostring());
                            if (pocet > 1) {
                                Target[i].PrurezMM2 = JedenKabel.SLmm2.ToString();
                                Target[i].PocetKabelu = pocet;
                            } 
                            else Target[i].PrurezMM2 = JedenKabel.SLmm2.ToString();
                            Target[i].Vodice = JedenKabel.Deleni;
                            Target[i].Kabel = JedenKabel;
                            volba = false;
                            break;
                        }
                        
                    }
                    //pocet++;
                }
            }
            return Target;
        }

        public static void DoplneniCsvToJson()
        {
            //Soubor kam bude doplněno
            string cestaData = Path.Combine(Cesty.ElektroDataJson);
            var Target = Soubory.LoadJsonList<Zarizeni>(cestaData);
            
            //Data pro doplnění
            string cesta = Path.Combine(Cesty.ElektroDataCsv);
            if (!File.Exists(cesta)) return;
            var Source = Soubory.LoadFromCsv<Zarizeni>(cesta);
            
            //Doplnění dat do Json.
            Prevod.UpdateCsvToJson(Source, Target );

            //uložit doplnění informace do Json
            Target.SaveJsonList(cestaData);
        }

        public static void VyvoritFMKM()
        {
            //var cesta = Environment.ProcessPath;            // Získá úplnou cestu ke spuštěnému procesu
            //var dir = Path.GetDirectoryName(cesta);         // Získá adresář, kde je spustitelný soubor
            VyvoritFM();
            VyvoritKM();
        }

        /// <summary> Převod seznamu frekvenčních měničů na Json </summary>
        public static void VyvoritFM()
        {
            string basePath = Path.Combine(Cesty.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv"); 
           
            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary> Převod seznamu stykačů na Json </summary>
        public static void VyvoritKM()
        {
            string basePath = Path.Combine(Cesty.BasePath, "Data");
            string CestaKM = Path.Combine(basePath, "KM.csv"); 
           
            var KM = Soubory.LoadFromCsv<Stykac>(CestaKM);
            Console.WriteLine($"Pocet stykaču: {KM.Count}");

            KM.SaveJsonList(Path.ChangeExtension(CestaKM, ".json"));
            Console.WriteLine($"Stykače uloženy jako Json");
        }

        /// <summary> Převod seznamu motorů a motorů3000 na jeden Json </summary>
        public static void VyvoritMotor()
        {
            string basePath = Path.Combine(Cesty.Data, "Motory");
            string CestaMotor = Path.Combine(basePath, "Motory.csv");
            var Motor = Soubory.LoadFromCsv<Motor>(CestaMotor);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            string CestaMotor3000 = Path.Combine(basePath, "Motory3000.csv");
            var Motor3000 = Soubory.LoadFromCsv<Motor>(CestaMotor3000);
            Console.WriteLine($"Pocet motorů: {Motor3000.Count}");

            Motor.AddRange(Motor3000);
            Console.WriteLine($"Pocet motorů: {Motor.Count}");

            Motor.SaveJsonList(Path.ChangeExtension(CestaMotor, ".json"));
            Console.WriteLine($"Motory uloženy jako Json");
        }

        public static void Rozvadec() {
            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            var Vývody = Path.Combine(Cesty.Elektro, "Vývody.json");
            var Data2 = Soubory.LoadJsonList<Zarizeni>(Vývody);

            Data = [.. Data, .. Data2];

            //rozdělení podle rozvaděče
            var skupiny = Data.GroupBy(x => x.RozvadecOznačení);
            int PocetRozvadecu = skupiny.Count();
            Console.WriteLine($"\nPočet rozvaděčů : {PocetRozvadecu}");

            foreach (var skupina in skupiny) {
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
                foreach(var item in Pole) {
                    //Console.WriteLine($"Rozvaděč: {skupina.Key}, Tag: {item.Tag}, Popis: {item.Popis}");
                    if(item.DruhEnum == Druhy.Přívod)
                        Console.WriteLine($"Tag: {item.Tag.Replace("\n"," "),-12}, Druh: {item.Druh,-12}, Popis: {item.Popis,-35}, SumaPříkon: {SumaPrikon,-15:F2}");
                    else
                        Console.WriteLine($"Tag: {item.Tag.Replace("\n", " "),-12}, Druh: {item.Druh,-12}, Popis: {item.Popis,-35}, Příkon: {item.Prikon,-15}");
                }

            }
        }
        public static void SpojitSeznamy()
        {
            var Data = Soubory.LoadJsonList<Zarizeni>(Cesty.ElektroDataJson);
            Data = [.. Data.Where(x => x.Etapa == "FAZE 1")];

            string cesta1 = Path.Combine(Cesty.Elektro, "Pid", @"UpravaZnovu.006.json");
            var Data2 = Soubory.LoadJsonList<Zarizeni>(Path.ChangeExtension(cesta1, ".json"));
            Data2 = [.. Data2.Where(x => x.Etapa == "FAZE 2")];

            Data = [.. Data, .. Data2];
            //Data.SaveJsonList(Path.Combine(Cesty.Elektro, "Pid", @"Test.json" ));
            //Data.SaveToCsv(Path.Combine(Cesty.Elektro, "Pid", @"Test.csv"));
            Data.SaveJsonList(Cesty.ElektroDataJson);
        }
        /// <summary>Převod stringu na enum</summary>
        private static void StringToEnum(IGrouping<string, Zarizeni> skupina) {
            foreach(var ukol in skupina) {
                if(Enum.TryParse<Druhy>(ukol.Druh, true, out var priorita)) {
                    ukol.DruhEnum = priorita;
                }
                else {
                    ukol.DruhEnum = Druhy.Nic; // nebo jiná výchozí hodnota
                }
            }
        }

        internal static void Duplicity()
        {
            string cestaData = Cesty.ElektroDataJson;
            var Data = Soubory.LoadJsonList<Zarizeni>(cestaData);
            Console.WriteLine($"Pocet záznamů: {Data.Count}");
            Data = [.. Data.DistinctBy(x => x.Apid)];

            //Verze přepsání původního Jsonu
            Data.SaveJsonList(cestaData);
        }

        internal static void NačtiSeznamVýkresůXls(string v) {
            var data = ExcelLoad.DataExcelVykres(v, "seznam dokumentace", 17);
            data.SaveJsonList(Path.ChangeExtension(v, ".json"));
        }
    }
}