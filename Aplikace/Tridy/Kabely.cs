using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Trasa
    {
        private string tag = string.Empty;

        public string Tag { get => tag; set => tag = value.Replace("\n", ""); } /// <summary>Jméno zařízení</summary>
        public string Rozvadec { get; set; } = string.Empty; //Zarizeni.Rozvadec
        public string RozvadecCislo { get; set; } = string.Empty; //Zarizeni.RozvadecCislo
        public string Oznaceni { get; set; } = string.Empty;    //"WL 01"

        public string Kabel { get; set; } = string.Empty;   //ozvaděčení kabelu
        public string PocetZil { get; set; } = string.Empty; //Zarizeni.vodice
        public string Prurezmm2 { get; set; } = string.Empty;   //Zarizeni.PruzezMM2
        public string PrurezFt { get; set; } = string.Empty; //nepoužito

        public string Druh { get; set; } = string.Empty;

        //Opakovani Tag
        public string OdkudSvokra { get; set; } = string.Empty;
        public string Mezera { get; set; } = string.Empty;
        public string Patro { get; set; } = string.Empty;
        public string Predmet { get; set; } = string.Empty;
        public string Svorka { get; set; } = string.Empty;

        public string Delka { get; set; } = string.Empty;
        /// <summary>Rozvaděč</summary>

        //převod enumu na pole stringů 
        public static string[] KabelZnačkaPole => Enum.GetNames<KabelZnačka>();
    }

    public class Kabely {
        public Trasa Hlavni { get; set; } 
        public Trasa PTC { get; set; }
        public Trasa Ovladani{ get; set; }
    }
    
    public enum KabelZnačka
    {
        WH,
        WL,
        WS,
        WC,
    }

    

    public class Kabel : Entity
    {
        //public string Deleni { get; set; } = string.Empty;
        public string Označení { get; set; } = string.Empty;
        
        //public string Proud { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;
        public string Proud { get; set; } = string.Empty;

        /// <summary>Počet vodičů</summary>
        [Display(Name = "Vodiče")]
        public string Deleni { get; set; } = string.Empty;

        public double SLmm2 { get; set; }

        //průřez PEN vodiče
        [Display(Name = "průřez PEN vodiče")]
        public double SPENmm2 { get; set; }

        /// <summary>proudové zatížení ve vzduchu svisle</summary>
        public double IzAGsvis { get; set; }

        /// <summary>proudové zatížení ve vzduchu vovorovně </summary>
        public double IzAGvod { get; set; }

        /// <summary>proudové zatížení ve vzduchu vedle sebe</summary>
        public double IzAFlin { get; set; }

        /// <summary>proudové zatížení ve  vzduchu trojůhelnik</summary>
        public double IzAFtroj { get; set; }

        //proudové zatížení v zemi
        public double IzAE { get; set; }

        //proudové zatížení v trubce v zemi
        public double IzAD1 { get; set; }

        //proudové zatížení přímo v zemi
        public double IzAD2 { get; set; }

        //proudové zatížení na stěně
        public double IzAC { get; set; }

        //proudové zatížení v trubce na stěně
        public double IzAB { get; set; }

        //proudové zatížení v izolační stěne
        public double IzAA { get; set; }

        //odpor krajního vodiče
        public double RLOhmkm { get; set; }

        //odpor PEN vodiče
        public double RPENOhmkm { get; set; }

        //induktance krajního vodiče
        public double XLOhmkm { get; set; }

        //induktance PEN vodiče
        public double XPENOhmkm { get; set; }

        //tau časová oteplovací konstanta vedení
        public double Taus { get; set; }

        public double TpracstC { get; set; }
        public double TpretstC { get; set; }
        public double TzkratstC { get; set; }

        //složky netočivé impedance vedení / složky sousledné impedance vedení
        public double RoR1 { get; set; }

        //složky netočivé impedance vedení / složky sousledné impedance vedení
        public double XoX1 { get; set; }
    }

    public class KabelVse : Kabel
    {
        public double MaxProud
        {
            get
            {
                double[] Poudy = [IzAGsvis, IzAGvod, IzAFlin, IzAFtroj, IzAE, IzAD1, IzAD2, IzAC, IzAB, IzAA];
                return Poudy.Max();
            }
        }

        public double MaxProudVzduch
        {
            get
            {
                double[] Poudy = [IzAGsvis, IzAGvod, IzAFlin, IzAFtroj];
                return Poudy.Max();
            }
        }


        //https://home.zcu.cz/~hejtman/PEC/Prednasky/pred4.pdf

        public static double DeltaU1f(KabelVse kabel, double proud, double delka, double uhel)
        {
            //Ubytek ve fazí
            var du = proud * ((kabel.RLOhmkm * Math.Cos(uhel)) + (kabel.XLOhmkm * Math.Sin(uhel)));
            //Ubytek ve Nule
            var duPen = proud * ((kabel.RPENOhmkm * Math.Cos(uhel)) + (kabel.XPENOhmkm * Math.Sin(uhel)));

            return (du + duPen) / 1000 * delka;
        }

        public static double DeltaU3f(KabelVse kabel, double proud, double delka, double uhel)
        {
            var odpor = (kabel.RLOhmkm * Math.Cos(uhel)) + (kabel.XLOhmkm * Math.Sin(uhel));
            var odporpe = (kabel.RPENOhmkm * Math.Cos(uhel)) + (kabel.XPENOhmkm * Math.Sin(uhel));
            var du = proud * (odpor + odporpe) / 1000;
            var v = delka * du; // / Math.Sqrt(3);
            return v;
        }

        public static double ProcentaU3f(KabelVse kabel, double napeti, double proud, double delka, double uhel)
        {
            return DeltaU3f(kabel, proud, delka, uhel) / napeti * 100;
        }
    }

    public static class Extension
    {
        public static double DeltaU3f(this KabelVse kabel, double proud, double delka, double uhel) =>
           KabelVse.DeltaU3f(kabel, proud, delka, uhel);

        public static double UProcenta(this KabelVse kabel, double napeti, double proud, double delka, double uhel) =>
            KabelVse.ProcentaU3f(kabel, napeti, proud, delka, uhel);
    }

}
