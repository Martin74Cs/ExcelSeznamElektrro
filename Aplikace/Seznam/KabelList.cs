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
        public List<List<string>> Kabely(List<List<string>> PoleData)
        {
            //uprava pole tabulky pro vypsaní
            var NovaData = new List<List<string>>();
            foreach (var radek in PoleData)
            {
                var Data = new List<string>();
                //Označení Kabel Odkud Kam

                //1. Kabel
                Data.Add("SO 01");
                //2. Kabel Tag
                Data.Add(radek[0]);
                //3. Kabel
                Data.Add("WL 01");
                //4. Jmeno kabelu
                if (radek[3] == "VSD")
                    Data.Add("ÖLFLEX CLASSIC 110 CY");
                else
                    Data.Add("CYKY");

                //5. Počet žil
                Data.Add("4x");

                //6. Průřez
                Data.Add(radek[9]);

                //7. Průřez
                Data.Add(radek[10]);

                //8. zařízení
                if (radek[3] == "PU")
                    Data.Add("Přívod");
                else
                    Data.Add("Motor");


                //9. odkud Mcc
                Data.Add(radek[11]);

                //10. Odkud číslo
                Data.Add(radek[12]);

                //11. Svorka
                Data.Add("X 01");

                //12. kam 
                Data.Add("SO 01");

                //13. kam číslo
                Data.Add(radek[0]);

                //14.kam Svorka
                Data.Add("X 01");

                //15. Delka m
                Data.Add(radek[7]);

                //16 Delka ft
                Data.Add(radek[8]);

                //přidání řádku do pole
                NovaData.Add(Data);

                //Ovládací kabel PTC
                if (radek[3] != "PU")
                { 
                  NovaData.Add(KabelPTC(radek));
                  NovaData.Add(KabelOvladani(radek));
                }



            }
            return NovaData;
        }

        public List<string> KabelPTC(List<string> radek)
        {
            //Ovládací kabel PTC
            var Data = new List<string>();

            //1. Kabel
            Data.Add("SO 01");

            //2. Kabel Tag
            Data.Add(radek[0]);

            //3. Kabel
            Data.Add("WS 01");

            //4. Kabel
            Data.Add("ÖLFLEX CLASSIC 100 ");

            //5. Počet žil
            Data.Add("2x");

            //6. Průřez mm2
            Data.Add(radek[9]);

            //7. Průřez awg
            Data.Add(radek[10]);

            //8. Počet žil
            Data.Add("Ptc");

            //9. odkud Mcc
            Data.Add(radek[11]);

            //10. Odkud číslo
            Data.Add(radek[12]);

            //11. Odkud Svorka
            Data.Add("X 01");



            //12. kam tag
            Data.Add(radek[0]);

            //13. kam číslo
            Data.Add("M 01");

            //14. kam Svorka
            Data.Add("X 01");



            //15. Delka m
            Data.Add(radek[7]);

            //16. Delka ft
            Data.Add(radek[8]);
                
            
            return Data;
        }

        public List<string> KabelOvladani(List<string> radek)
        {
            //Ovládací kabel PTC
            var Data = new List<string>();

            //1. Kabel
            Data.Add("SO 01");

            //2. Kabel Tag
            Data.Add(radek[0]);

            //3. Kabel
            Data.Add("WS 02");

            //4. Kabel
            Data.Add("CYKY");

            //5. Počet žil
            Data.Add("12x");

            //6. Průřez mm2
            Data.Add(radek[9]);

            //7. Průřez awg
            Data.Add(radek[10]);

            //8. Počet žil
            Data.Add("Ovládání");

            //9. odkud Mcc
            Data.Add(radek[11]);

            //10. Odkud číslo
            Data.Add(radek[12]);

            //11. Odkud Svorka
            Data.Add("X 01");



            //12. kam tag
            Data.Add(radek[0]);

            //13. kam číslo
            Data.Add("MX 01");

            //14. kam Svorka
            Data.Add("X 01");


            //15. Delka m
            Data.Add(radek[7]);

            //16. Delka ft
            Data.Add(radek[8]);


            return Data;
        }
    }
}
