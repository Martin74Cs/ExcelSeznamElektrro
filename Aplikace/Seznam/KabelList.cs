using Aplikace.Excel;
using Aplikace.Tridy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Aplikace.Seznam
{
    public class KabelList
    {
        public static List<List<string>> Kabely(List<Zarizeni> PoleData)
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
                Data.Add(radek.Delka);

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
                radek.Tag,

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
                "13",

                //9. Počet žil
                "Ptc",



                //10. odkud tag
                radek.Tag,

                //11. odkud Mcc
                radek.Rozvadec,

                //12. Odkud číslo
                radek.RozvadecCislo,

                //13. Odkud Svorka
                "X 02",

                //14. Mezera
                " ",


                //15. kam tag
                radek.Tag,

                //16. kam číslo
                "SO 01",

                //17. kam číslo
                "M 01",

                //18. kam Svorka
                "X 02",

                //19. Delka m
                radek.Delka,

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
                "13",

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
                "SO 01",

                //17. kam číslo
                "MX 01",

                //18. kam Svorka
                "X 01",

                //19. Delka m
                radek.Delka,

                //20. Delka ft
                //radek.Delka
            };


            return Data;
        }
    }
}
