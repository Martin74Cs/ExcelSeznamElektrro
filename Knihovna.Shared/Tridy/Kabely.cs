using Knihovna.Tridy;
using System.ComponentModel.DataAnnotations;

namespace Knihovna.Shared.Tridy
{
    /// <summary>
    /// Reprezentuje kabelovou trasu.
    /// </summary>
    public class Trasa : Entity
    {
        private string tag = string.Empty;

        public string Tag { get => tag; set => tag = value.Replace("\n", ""); } /// <summary>Jméno zařízení</summary>
        public string Rozvadec { get; set; } = string.Empty; //Zarizeni.Rozvadec
        public string RozvadecCislo { get; set; } = string.Empty; //Zarizeni.RozvadecCislo

        [Display(Name = "Rozváděč")]
        [Newtonsoft.Json.JsonIgnore]
        public string RozvadecAll => Rozvadec + " " + RozvadecCislo;

        public string Oznaceni { get; set; } = string.Empty;    //"WL 01"

        [Display(Name = "Kabel typ")]
        public string Kabel { get; set; } = string.Empty;   //ozvaděčení kabelu
        public string PocetZil { get; set; } = string.Empty; //Zarizeni.vodice
        public string Prurezmm2 { get; set; } = string.Empty;   //Zarizeni.PrurezMM2

        [Newtonsoft.Json.JsonIgnore]
        public string KabelAll => Kabel + " " + PocetZil + "x" + Prurezmm2;

        [Display(Name = "Průřez/Size")]
        [Newtonsoft.Json.JsonIgnore]
        public string KabelVelikost { get; set; } = string.Empty; // => PocetZil + "x" + Prurezmm2;

        public string PrurezFt { get; set; } = string.Empty; //nepoužito
        public string Druh { get; set; } = string.Empty;

        //Opakovani Tag

        [Display(Name = "Ukončení")]
        public string OdkudSvorka { get; set; } = string.Empty;
        public string Mezera { get; set; } = string.Empty;
        public string Patro { get; set; } = string.Empty;
        public string Predmet { get; set; } = string.Empty;

        [Display(Name = "Ukončení")]
        public string Svorka { get; set; } = string.Empty;

        [Display(Name = "Délka/Lengh")]
        public string Delka { get; set; } = string.Empty;
        public string Popis { get; set; } = string.Empty;
        /// <summary>Rozvaděč</summary>
        [System.ComponentModel.Category("5. Kabelové připojení")]
        [System.ComponentModel.DisplayName("Kabel (Objekt)")]
        [System.ComponentModel.Description("Interní data kabelu.")]
        public Kabel? KabelData { get; set; } = new Kabel();

        /// <summary>
        /// Vyhledá a aktualizuje interní data kabelu (proudové zatížení, odpor atd.) na základě typu, počtu žil a průřezu.
        /// </summary>
        public void AktualizujKabelData()
        {
            KabelData = KabelDatabaze.NajdiKabel(Kabel, PocetZil, Prurezmm2) ?? new Kabel();
        }

        //převod enumu na pole stringů 
    }

    /*
    /// <summary>
    /// Reprezentuje skupinu kabelů pro zařízení.
    /// </summary>
    public class Kabely
    {
        public Trasa Hlavni { get; set; } 
        public Trasa PTC { get; set; }
        public Trasa Ovladani{ get; set; }
    }
    */

    /*
    public enum KabelZnačka
    {
        WH,
        WL,
        WS,
        WC,
    }
    */

    

    public class Kabel : Entity
    {
        public string Označení { get; set; } = string.Empty;
        

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

        public static double VypoctiProudZatizeni(double prikonKw, double napetiV, double cosPhi)
        {
            if (prikonKw <= 0 || napetiV <= 0 || cosPhi <= 0) return 0;
            if (napetiV >= 380)
            {
                return (prikonKw * 1000.0) / (Math.Sqrt(3.0) * napetiV * cosPhi);
            }
            else
            {
                return (prikonKw * 1000.0) / (napetiV * cosPhi);
            }
        }

        public static double VypoctiUbytekNapetiV(this Kabel kabel, double proudA, double delkaM, double napetiV, double cosPhi)
        {
            if (kabel == null || proudA <= 0 || delkaM <= 0 || napetiV <= 0 || cosPhi <= 0) return 0;
            double sinPhi = Math.Sqrt(1.0 - cosPhi * cosPhi);
            
            // Přepočet odporu na maximální provozní teplotu kabelu (standardně 70 °C pro PVC, 90 °C pro XLPE)
            double tprac = kabel.TpracstC > 0 ? kabel.TpracstC : 70.0;
            double tempCoeff = 1.0 + 0.00393 * (tprac - 20.0);
            double rTemp = kabel.RLOhmkm * tempCoeff;
            
            if (napetiV >= 380)
            {
                return Math.Sqrt(3.0) * proudA * ((rTemp * cosPhi) + (kabel.XLOhmkm * sinPhi)) * (delkaM / 1000.0);
            }
            else
            {
                return 2.0 * proudA * ((rTemp * cosPhi) + (kabel.XLOhmkm * sinPhi)) * (delkaM / 1000.0);
            }
        }

        public static double VypoctiUbytekNapetiProcenta(this Kabel kabel, double proudA, double delkaM, double napetiV, double cosPhi)
        {
            if (napetiV <= 0) return 0;
            double ubytekV = kabel.VypoctiUbytekNapetiV(proudA, delkaM, napetiV, cosPhi);
            return (ubytekV / napetiV) * 100.0;
        }

        public static double VypoctiTeplotuVodice(this Kabel kabel, double proudA, bool veVzduchu)
        {
            if (kabel == null || proudA <= 0) return 30.0;
            double tprac = kabel.TpracstC > 0 ? kabel.TpracstC : 70.0;
            double iz = veVzduchu ? kabel.IzAGvod : kabel.IzAC;
            if (iz <= 0) return 30.0;
            
            return 30.0 + (tprac - 30.0) * Math.Pow(proudA / iz, 2.0);
        }

        public static double VypoctiZkratovyProud(this Kabel kabel, bool jeHlinik, double casSekund)
        {
            if (kabel == null || casSekund <= 0) return 0;
            double k = jeHlinik ? 76.0 : 115.0;
            double s = kabel.SLmm2;
            return (k * s) / Math.Sqrt(casSekund);
        }

        public static double VypoctiImpedanciSmycky(this Kabel kabel, double delkaM)
        {
            if (kabel == null || delkaM <= 0) return 0;
            double tprac = kabel.TpracstC > 0 ? kabel.TpracstC : 70.0;
            double tempCoeff = 1.0 + 0.00393 * (tprac - 20.0);
            
            double rLoop = (kabel.RLOhmkm + kabel.RPENOhmkm) * tempCoeff * (delkaM / 1000.0);
            double xLoop = (kabel.XLOhmkm + kabel.XPENOhmkm) * (delkaM / 1000.0);
            return Math.Sqrt(rLoop * rLoop + xLoop * xLoop);
        }

        public static double VypoctiZkrat3f(this Kabel kabel, double delkaM, double napetiV)
        {
            if (kabel == null || delkaM <= 0 || napetiV <= 0) return 0;
            double rL = kabel.RLOhmkm * (delkaM / 1000.0);
            double xL = kabel.XLOhmkm * (delkaM / 1000.0);
            double zL = Math.Sqrt(rL * rL + xL * xL);
            if (zL <= 0) return 0;
            return napetiV / (Math.Sqrt(3.0) * zL);
        }

        public static double VypoctiZkrat1f(this Kabel kabel, double delkaM, double napetiV)
        {
            if (kabel == null || delkaM <= 0 || napetiV <= 0) return 0;
            double rLoop = (kabel.RLOhmkm + kabel.RPENOhmkm) * (delkaM / 1000.0);
            double xLoop = (kabel.XLOhmkm + kabel.XPENOhmkm) * (delkaM / 1000.0);
            double zLoop = Math.Sqrt(rLoop * rLoop + xLoop * xLoop);
            if (zLoop <= 0) return 0;
            
            double uPhase = napetiV / Math.Sqrt(3.0);
            return uPhase / zLoop;
        }
    }
}

