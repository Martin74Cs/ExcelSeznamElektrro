using Aplikace.Excel;
using Aplikace.Tridy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Aplikace.Seznam
{
    public class KabelList
    {
        public static List<List<string>> KabelyOld(List<Zarizeni> PoleData)
        {
            //string[] TextPole =     ["Tag", "PID", "Popis", "Prikon", "BalenaJednotka", "Menic", "Proud500", "HP", "Proud480", "mm2", "AWG", "Delkam", "Delkaft", "MCC", "cisloMCC"];
            //int[] PouzitProTabulku1 = [3,   2,      7,      18,         1,              21,         59,     56,     60,         63,     64,     61,     62,         65,     66];

            //Stara.Add(["P101", "V101", "Čeradlo", "10", "", "VSD", "", "100", "MCC", "1",]);

            //uprava pole tabulky pro vypsaní
            var NovaData = new List<List<string>>();
            foreach (var radek in PoleData)
            {
                var Data = new List<string>
                {
                    //1. Kabel
                    radek.Tag,

                    //2. odkud Mcc
                    radek.Rozvadec,

                    //3. Odkud číslo
                    radek.RozvadecCislo,

                    //4. Kabel
                    "WL 01"
                };

                //5. Jmeno kabelu
                if (radek.Menic == "VSD")
                    Data.Add("ÖLFLEX CLASSIC 110 CY");
                else
                    Data.Add("CYKY");

                //6. Počet žil
                if (radek.Menic == "VSD")
                    Data.Add("4x");
                else
                    Data.Add("5x");             

                //7. proud
                Data.Add(radek.PruzezMM2);

                //8. Průřez
                Data.Add(radek.PruzezMM2);

                //9. zařízení
                //Pokud začíná na P ne B jedná se  balenou jednotku
                if (radek.BalenaJednotka.StartsWith('P') || radek.BalenaJednotka.StartsWith('B'))
                    Data.Add("Přívod");
                else
                    Data.Add("Motor");

                //10. odkud tag
                Data.Add(radek.Tag);

                //11. odkud Mcc
                Data.Add(radek.Rozvadec);

                //12. Odkud číslo
                Data.Add(radek.RozvadecCislo);

                //13. Svorka rozvaděče
                Data.Add("X 01");

                //14. Mezera
                Data.Add(" ");




                //15. kam tag
                Data.Add(radek.Tag);

                //16. kam objekt nebo patro
                Data.Add("SO 01");

                //17.kam Zažizeni
                //18.kam Svorka
                if (radek.BalenaJednotka.StartsWith('P') || radek.BalenaJednotka.StartsWith('B'))
                {
                    Data.Add(radek.BalenaJednotka);
                    Data.Add("X 01");
                }
                else 
                { 
                    Data.Add("M 01");
                    Data.Add("X 01");
                }

                //19. Delka m
                Data.Add(radek.Delka.ToString());

                //20 Delka ft
                //Data.Add(radek.Delka);

                //přidání řádku do pole
                NovaData.Add(Data);

                //Ovládací kabel PTC
                if (!radek.BalenaJednotka.StartsWith('P') && !radek.BalenaJednotka.StartsWith('B'))
                { 
                  NovaData.Add(KabelPTC(radek));
                  NovaData.Add(KabelOvladani(radek));
                }
            }
            return NovaData;
        }
        
        public static List<List<string>> Kabely(List<Zarizeni> PoleData)
        {
            //uprava pole tabulky pro vypsaní
            var NovaData = new List<List<string>>();
            Trasa trasa = new(); //použití nové třídy pro trasy
            foreach (var radek in PoleData)
            {
                var Tag = radek.Tag.Replace("\n", " "); //1. Kabel
                var Rozvadec = radek.Rozvadec; //2. odkud Mcc
                var Cislo = radek.RozvadecCislo; //3. Odkud číslo
                string Oznaceni = "WL 01"; //4. Kabel

                trasa.Tag = radek.Tag.Replace("\n", " "); 
                trasa.Rozvadec = Rozvadec;
                trasa.RozvadecCislo = Cislo;
                trasa.Oznaceni = Cislo;
                
                var Data = new List<string>  {
                    Tag,        //1. Kabel
                    Rozvadec,   //2. odkud Mcc
                    Cislo,      //3. Odkud číslo
                    Oznaceni,   //4. Kabel
                };

                string Kabel;    //5. Jmeno kabelu
                string PocetZil; //6. Počet žil
                if (radek.Menic == "VSD")
                {
                    Kabel = "ÖLFLEX CLASSIC 110 CY";
                    PocetZil = "4x";
                }
                else { 
                    Kabel = radek.Kabel.Označení ?? "";
                    PocetZil = "5x";
                }
                string PruzezMM2 = radek.PruzezMM2; //7. Průřez
                string PrurezFt = ""; //8. 

                trasa.Kabel = Kabel;
                trasa.PocetZil = PocetZil;
                trasa.Prurezmm2 = PruzezMM2;
                trasa.PrurezFt = PrurezFt;

                var Data2 = new List<string>  {
                    Kabel,      //5.
                    PocetZil,   //6.
                    PruzezMM2,  //7.
                    PrurezFt,   //8.
                };
                Data.AddRange(Data2);

                trasa.Druh = radek.Druh; //9. zařízení
                Data.Add(radek.Druh);

                //10.11.12.13
                //dtto trasa - trasa.Tag; 
                Data.Add(radek.Tag);    //10. odkud tag
                //dtto trasa - trasa.Rozvadec; 
                Data.Add(trasa.Rozvadec);   //11. odkud Mcc
                //dtto trasa - trasa.RozvadecCislo; 
                Data.Add(radek.RozvadecCislo);  //12. Odkud číslo
                //13. Svorka rozvaděče
                trasa.OdkudSvokra = "X 01"; 
                Data.Add(trasa.OdkudSvokra);

                //14. Mezera
                trasa.Mezera = "";
                Data.Add(trasa.Mezera);

                //15.16.17.18
                //15. kam tag
                Data.Add(trasa.Tag);

                //16. kam objekt nebo patro
                Data.Add(radek.Patro);

                //17.kam Zažizeni
                //18.kam Svorka
                Data.Add(radek.Predmet);
                Data.Add("X 01");

                //19. Delka m
                Data.Add(radek.Delka.ToString());

                //20 Delka ft
                //Data.Add(radek.Delka);

                //přidání řádku do pole
                NovaData.Add(Data);

                //Ovládací kabel PTC
                if (!radek.BalenaJednotka.StartsWith('P') && !radek.BalenaJednotka.StartsWith('B'))
                { 
                  NovaData.Add(KabelPTC(radek));
                  NovaData.Add(KabelOvladani(radek));
                }
            }
            return NovaData;
        }
        public static List<string> KabelPTC(Zarizeni radek)
        {
            //Ovládací kabel PTC
            var Data = new List<string>
            {
                //1. Kabel tag
                radek.Tag.Replace("\n", " "),

                //2. Kabel MCC
                radek.Rozvadec,

                //3. Kabel MCC cislo
                radek.RozvadecCislo,

                //4. Kabel
                "WS 01",


                //5. Kabel
                "ÖLFLEX CLASSIC 100 ",

                //6. Počet žil
                "2x",

                //7. Průřez mm2
                "2,5",

                //8. Průřez awg
                //"13",
                "",

                //9. Počet žil
                "Ptc",

                //10. odkud tag
                radek.Tag.Replace("\n", " "),

                //11. odkud Mcc
                radek.Rozvadec,

                //12. Odkud číslo
                radek.RozvadecCislo,

                //13. Odkud Svorka
                "X 02",

                //14. Mezera
                " ",


                //15. kam tag
                radek.Tag.Replace("\n", " "),

                //16. kam číslo
                "Patro",

                //17. kam číslo
                "M 01",

                //18. kam Svorka
                "X 02",

                //19. Delka m
                radek.Delka.ToString(),

                //20. Delka ft
                //radek.Delka
            };
                
            
            return Data;
        }
        public static List<string> KabelOvladani(Zarizeni radek)
        {
            //Ovládací kabel PTC
            var Data = new List<string>
            {
                //1. Kabel tag
                radek.Tag,

                //2. odkud Mcc
                radek.Rozvadec,

                //3. Odkud číslo
                radek.RozvadecCislo,

                //4. Kabel
                "WS 02",



                //5. Kabel
                "CYKY",

                //6. Počet žil
                "12x",

                //7. Průřez mm2
                "2,5",

                //8. Průřez awg
                //"13",
                "",

                //9. Počet žil
                "Ovládání",



                //10. odkud tag
                radek.Tag,

                //11. odkud Mcc
                radek.Rozvadec,

                //12. Odkud číslo
                radek.RozvadecCislo,

                //13. Odkud Svorka
                "X 03",

                //14. Mezera
                " ",



                //15. kam tag
                radek.Tag,

                //16. kam číslo
                "Patro",

                //17. kam číslo
                "MX 01",

                //18. kam Svorka
                "X 01",

                //19. Delka m
                radek.Delka.ToString(),

                //20. Delka ft
                //radek.Delka
            };
            return Data;
        }

        /// <summary>Převod ListTrasa do Listu rádku a  Listu sloupců pro zadaní do excel </summary>
        public static List<List<string>> KabelyTridaToString(List<Trasa> Data)
        {
            //uprava pole tabulky pro vypsaní
            var dir = new Dictionary<int, string>() {
                //Označnení kaeblu
                {1,"Tag"},
                {2,"Rozvadec"},
                {3,"RozvadecCislo"},
                {4,"Oznaceni"},
                //Druh kabel
                {5,"Kabel"},
                {6,"PocetZil"},
                {7,"Prurezmm2"},
                {8,"PrurezFt"},
                //druh
                {9,"Druh"},
                //Kabel odkud
                {10,"Tag"},
                {11,"Rozvadec"},
                {12,"RozvadecCislo"},
                {13,"OdkudSvokra"},
                //mezera
                {14,"Nic"},
                //Kabel Kam
                {15,"Tag"},
                {16,"Predmet"},
                {17,"Patro"},
                {18,"Svorka"},

                {19,"Delka"},

            };
            var Pole = new List<List<string>>();
            foreach (var item in Data)
            {
                var result = dir.OrderBy(kvp => kvp.Key) // Seřadíme podle int klíče
                    .Select(kvp => {
                        var prop = item.GetType().GetProperty(kvp.Value);
                        return prop != null ? prop.GetValue(item)?.ToString() ?? "" : "";
                    })
                    .ToList();
                Pole.Add(result);
            }
            return Pole;
        }


        /// <summary> Vytvoření pole kabelů pro zápis do Excelu </summary>
        // public static List<Trasa> KabelyTrida(List<Zarizeni> PoleData)
        public static List<Kabely> KabelyTrida(List<Zarizeni> PoleData)
        {
            var Kabely = new List<Kabely>();
            //uprava pole tabulky pro vypsaní
            var NovaData = new List<Trasa>();
            foreach (var radek in PoleData)
            {
                var Kabel = new Kabely();
                Trasa trasa = new();
                //1.2.3.4
                trasa.Tag = radek.Tag.Replace("\n", " "); //1. Kabel
                trasa.Rozvadec = radek.Rozvadec;          //2. odkud Mcc
                trasa.RozvadecCislo = radek.RozvadecCislo;//3. Odkud číslo
                trasa.Oznaceni = "WL 01";                 //4. Kabel

                //5.6.
                if (radek.Menic == "VSD")
                {
                    trasa.Kabel = "ÖLFLEX CLASSIC 110 CY";  //5. Kabel
                    trasa.PocetZil = "4x";                  //6. Kabel PocetZil
                }
                else { 
                    //trasa.Kabel = radek.Kabel.Označení ?? "";  //5. Kabel
                    trasa.Kabel = "PraflaDur";  //5. Kabel
                    trasa.PocetZil = "5x";                     //6. Kabel PocetZil
                }
                //7.8
                trasa.Prurezmm2 = radek.PruzezMM2; //7. Průřez
                //trasa.PrurezFt = "";               //8. Prozatím nepoužito

                //9. zařízení
                //if (string.IsNullOrEmpty(radek.Druh))
                //{
                //    if (radek.Druh == Zarizeni.Druhy.Rozvadeč.ToString())
                //    //if (radek.BalenaJednotka.StartsWith('P') || radek.BalenaJednotka.StartsWith('B') || radek.BalenaJednotka.StartsWith('x'))
                //        trasa.Druh = "Odhad" + "Motor"; //9. zařízení
                //    else
                //        trasa.Druh = "Odhad" + "Rozvaděč"; //9. zařízení
                //}
                //else
                trasa.Druh = radek.Druh; //9. zařízení

                //10.11.12.13
                //trasa.Tag = trasa.Tag;             //10. odkud tag
                //trasa.Rozvadec = trasa.Rozvadec;   //11. odkud Mcc
                //trasa.RozvadecCislo = trasa.RozvadecCislo; //12. Odkud číslo
                trasa.OdkudSvokra = "X 01";         //13. Svorka rozvaděče

                trasa.Mezera = "";                //14. Mezera

                //15.16.17.18
                //trasa.Tag = radek.Tag;          //15. kam tag
                trasa.Patro = radek.Patro;      //16. kam objekt nebo patro
                trasa.Predmet = radek.Predmet;  //17.kam Zažizeni
                trasa.Svorka = "X 01";          //18.kam Svorka
                
                trasa.Delka = (radek.Delka/1000).ToString("F2");      //19. Delka m
                //trasa.PrurezFt = (radek.Delka * (1/0.3)).ToString(); //20 Delka ft

                //přidání řádku do pole
                NovaData.Add(trasa);
                Kabel.Hlavni = trasa; //přiřazení hlavní trasy do kabelu
                                      //
                                      //Ovládací kabel PTC
                //if (!radek.BalenaJednotka.StartsWith('P') && !radek.BalenaJednotka.StartsWith('B'))
                if (radek.Druh == Zarizeni.Druhy.Motor.ToString())
                {
                    var PTC = KabelPTCTrida(radek);
                    NovaData.Add(PTC);
                    Kabel.PTC = PTC; //přiřazení PTC trasy do kabelu

                    var Ovladani = KabelOvladaniTrida(radek);
                    NovaData.Add(Ovladani);
                    Kabel.Ovladani = Ovladani; //přiřazení ovládací trasy do kabelu
                }
                Kabely.Add(Kabel); //přidání kabelu do seznamu kabelů   
            }
            //return NovaData;
            return Kabely;
        }

        public static Trasa KabelPTCTrida(Zarizeni radek)
        {
            //Ovládací kabel PTC
            Trasa trasa = new(); //Použití nové třídy pro trasy

            trasa.Tag = radek.Tag.Replace("\n", " "); //1. Kabel
            trasa.Rozvadec = radek.Rozvadec;          //2. odkud Mcc
            trasa.RozvadecCislo = radek.RozvadecCislo;//3. Odkud číslo
            trasa.Oznaceni = "WS 01";                 //4. Kabel

            trasa.Kabel = "ÖLFLEX CLASSIC 100 ";    //5. Kabel
            trasa.PocetZil = "2x";                  //6. Kabel PocetZil
            trasa.Prurezmm2 = "2,5";                //7. Průřez
            //trasa.PrurezFt = "";                    //8. Prozatím nepoužito

            trasa.Druh = "Ptc";                    //9. Druh

            //Odkud
            //Tag                   10
            //Rozvadec              11
            //RozvadecCislo         12
            trasa.OdkudSvokra = "X 02"; //13. Svorka rozvaděče
           
            trasa.Mezera = "";                //14. Mezera

            //Kam 15.16.17.18
            //trasa.Tag = radek.Tag;          //15. kam tag
            trasa.Patro = radek.Patro;      //16. kam objekt nebo patro
            trasa.Predmet = radek.Predmet;  //17.kam Zažizeni
            trasa.Svorka = "X 01";          //18.kam Svorka

            trasa.Delka = (radek.Delka / 1000).ToString("F2");                 //19. Delka m
            
            return trasa;
        }
        public static Trasa KabelOvladaniTrida(Zarizeni radek)
        {
            Trasa trasa = new(); //Použití nové třídy pro trasy

            trasa.Tag = radek.Tag.Replace("\n", " "); //1. Kabel
            trasa.Rozvadec = radek.Rozvadec;          //2. odkud Mcc
            trasa.RozvadecCislo = radek.RozvadecCislo;//3. Odkud číslo
            trasa.Oznaceni = "WS 02";                 //4. Kabel

            trasa.Kabel = "PRAFlaDur ";    //5. Kabel
            trasa.PocetZil = "12x";                  //6. Kabel PocetZil
            trasa.Prurezmm2 = "2,5 RE";                //7. Průřez
            trasa.PrurezFt = "";                    //8. Prozatím nepoužito

            trasa.Druh = "Ovládání";                    //9. Druh

            //Odkud
            //Tag                   10
            //Rozvadec              11
            //RozvadecCislo         12
            trasa.OdkudSvokra = "X 03"; //13. Svorka rozvaděče

            trasa.Mezera = "";                //14. Mezera

            //Kam 15.16.17.18
            //trasa.Tag = radek.Tag;        //15. kam tag
            trasa.Patro = radek.Patro;    //16. kam objekt nebo patro
            trasa.Predmet = "MX 01";      //17.kam Zažizeni
            trasa.Svorka = "X 01";        //18.kam Svorka

            trasa.Delka = (radek.Delka / 1000).ToString("F2");                 //19. Delka m

            return trasa;
        }
    }
}
