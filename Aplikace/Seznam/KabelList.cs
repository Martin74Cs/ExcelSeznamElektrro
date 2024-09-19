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

                //1. Kabel
                Data.Add(radek[0]);

                 //2. odkud Mcc
                Data.Add(radek[11]);

                //3. Odkud číslo
                Data.Add(radek[12]);

                //4. Kabel
                Data.Add("WL 01");



                //5. Jmeno kabelu
                if (radek[3] == "VSD")
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
                if (radek[3] == "PU")
                    Data.Add("Přívod");
                else
                    Data.Add("Motor");



                //10. odkud tag
                Data.Add(radek[0]);

                //11. odkud Mcc
                Data.Add(radek[11]);

                //12. Odkud číslo
                Data.Add(radek[12]);

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
                if (radek[3] == "PU")
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
                Data.Add(radek[7]);

                //20 Delka ft
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

            //1. Kabel tag
            Data.Add(radek[0]);

            //2. Kabel MCC
            Data.Add(radek[11]);

            //3. Kabel MCC cislo
            Data.Add(radek[12]);

            //4. Kabel
            Data.Add("WS 01");



            //5. Kabel
            Data.Add("ÖLFLEX CLASSIC 100 ");

            //6. Počet žil
            Data.Add("2x");

            //7. Průřez mm2
            Data.Add("2,5");

            //8. Průřez awg
            Data.Add("13");

            //9. Počet žil
            Data.Add("Ptc");



            //10. odkud tag
            Data.Add(radek[0]);

            //11. odkud Mcc
            Data.Add(radek[11]);

            //12. Odkud číslo
            Data.Add(radek[12]);

            //13. Odkud Svorka
            Data.Add("X 02");

            //14. Mezera
            Data.Add(" ");


            //15. kam tag
            Data.Add(radek[0]);

            //16. kam číslo
            Data.Add("SO 01");

            //17. kam číslo
            Data.Add("M 01");

            //18. kam Svorka
            Data.Add("X 02");



            //19. Delka m
            Data.Add(radek[7]);

            //20. Delka ft
            Data.Add(radek[8]);
                
            
            return Data;
        }

        public List<string> KabelOvladani(List<string> radek)
        {
            //Ovládací kabel PTC
            var Data = new List<string>();

            //1. Kabel tag
            Data.Add(radek[0]);
            
            //2. odkud Mcc
            Data.Add(radek[11]);

            //3. Odkud číslo
            Data.Add(radek[12]);

            //4. Kabel
            Data.Add("WS 02");



            //5. Kabel
            Data.Add("CYKY");

            //6. Počet žil
            Data.Add("12x");

            //7. Průřez mm2
            Data.Add("2,5");

            //8. Průřez awg
            Data.Add("13");

            //9. Počet žil
            Data.Add("Ovládání");



            //10. odkud tag
            Data.Add(radek[0]);

            //11. odkud Mcc
            Data.Add(radek[11]);

            //12. Odkud číslo
            Data.Add(radek[12]);

            //13. Odkud Svorka
            Data.Add("X 03");

            //14. Mezera
            Data.Add(" ");



            //15. kam tag
            Data.Add(radek[0]);

            //16. kam číslo
            Data.Add("SO 01");

            //17. kam číslo
            Data.Add("MX 01");

            //18. kam Svorka
            Data.Add("X 01");




            //19. Delka m
            Data.Add(radek[7]);

            //20. Delka ft
            Data.Add(radek[8]);


            return Data;
        }
    }
}
