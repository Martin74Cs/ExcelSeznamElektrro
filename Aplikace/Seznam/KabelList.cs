using Aplikace.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Aplikace.Seznam
{
    public class KabelList
    {
        public static List<List<string>> Kabely(List<List<string>> PoleData)
        {
            //uprava pole tabulky pro vypsaní
            var NovaData = new List<List<string>>();
            foreach (var radek in PoleData)
            {
                var Data = new List<string>
                {
                    //1. Kabel
                    radek[0],

                    //2. odkud Mcc
                    radek[13],

                    //3. Odkud číslo
                    radek[14],

                    //4. Kabel
                    "WL 01"
                };



                //5. Jmeno kabelu
                if (radek[5] == "VSD")
                    Data.Add("ÖLFLEX CLASSIC 110 CY");
                else
                    Data.Add("CYKY");

                //6. Počet žil
                Data.Add("4x");

                //7. Průřez
                Data.Add(radek[9]);

                //8. Průřez
                Data.Add(radek[10]);




                //9. zařízení
                if (radek[4] == "PU")
                    Data.Add("Přívod");
                else
                    Data.Add("Motor");



                //10. odkud tag
                Data.Add(radek[0]);

                //11. odkud Mcc
                Data.Add(radek[13]);

                //12. Odkud číslo
                Data.Add(radek[14]);

                //13. Svorka rozvaděče
                Data.Add("X 01");

                //14. Mezera
                Data.Add(" ");




                //15. kam tag
                Data.Add(radek[0]);

                //16. kam objekt nebo patro
                Data.Add("SO 01");

                //17.kam Zažizeni
                //18.kam Svorka
                if (radek[4] == "PU")
                {
                    Data.Add("PU 01");
                    Data.Add("X 01");
                }
                else 
                { 
                    Data.Add("M 01");
                    Data.Add("X 01");
                }



                //19. Delka m
                Data.Add(radek[11]);

                //20 Delka ft
                Data.Add(radek[12]);

                //přidání řádku do pole
                NovaData.Add(Data);

                //Ovládací kabel PTC
                if (radek[4] != "PU")
                { 
                  NovaData.Add(KabelPTC(radek));
                  NovaData.Add(KabelOvladani(radek));
                }



            }
            return NovaData;
        }

        public static List<string> KabelPTC(List<string> radek)
        {
            //Ovládací kabel PTC
            var Data = new List<string>
            {
                //1. Kabel tag
                radek[0],

                //2. Kabel MCC
                radek[13],

                //3. Kabel MCC cislo
                radek[14],

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
                radek[0],

                //11. odkud Mcc
                radek[13],

                //12. Odkud číslo
                radek[14],

                //13. Odkud Svorka
                "X 02",

                //14. Mezera
                " ",


                //15. kam tag
                radek[0],

                //16. kam číslo
                "SO 01",

                //17. kam číslo
                "M 01",

                //18. kam Svorka
                "X 02",



                //19. Delka m
                radek[11],

                //20. Delka ft
                radek[12]
            };
                
            
            return Data;
        }

        public static List<string> KabelOvladani(List<string> radek)
        {
            //Ovládací kabel PTC
            var Data = new List<string>
            {
                //1. Kabel tag
                radek[0],

                //2. odkud Mcc
                radek[13],

                //3. Odkud číslo
                radek[14],

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
                radek[0],

                //11. odkud Mcc
                radek[13],

                //12. Odkud číslo
                radek[14],

                //13. Odkud Svorka
                "X 03",

                //14. Mezera
                " ",



                //15. kam tag
                radek[0],

                //16. kam číslo
                "SO 01",

                //17. kam číslo
                "MX 01",

                //18. kam Svorka
                "X 01",




                //19. Delka m
                radek[11],

                //20. Delka ft
                radek[12]
            };


            return Data;
        }
    }
}
