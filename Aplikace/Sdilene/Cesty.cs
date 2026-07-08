using Aplikace.Tridy;
using Knihovna.Tridy;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Sdilene
{
    public static class Cesty
    {

        ///<summary> soubor spuštení exe </summary>
        public static string SouborExe => System.Reflection.Assembly.GetExecutingAssembly().Location;

        ///<summary> adresar spušteni dle souboru exe</summary>
        public static string AdresarSpusteni => System.IO.Path.GetDirectoryName(SouborExe);

        /// <summary> Cesta ProgramFiles</summary>
        public static string ProgramFiles { get => Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles); }

        public static string UserProfile { get => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile); }
        public static string Bundle { get => Path.Combine(UserProfile, @"AppData\Roaming\Autodesk\ApplicationPlugins\Elektro.bundle"); }
        public static string SichrAcad => Path.Combine(Bundle, "Sichr");
        public static string CuJsonAcad => Path.Combine(SichrAcad, "Cu.Json");
        public static string AlJsonAcad => SichrAcad + @"Al.Json";

        /// <summary>...LightChem\Elektro\Lightchem </summary>

        /// <summary>...Můj disk\Elektro\Lightchem\Ëlektro </summary>
        //public static string Elektro { get {
        //        //if (!Directory.Exists(Cesty.Elektro))
        //    }
        //}

        
        //public static string BasePath {
        //    get {
        //        //if (Environment.UserDomainName == "D10")
        //        //else

        //            OpenFileDialog dialog = new() {
        //                Title = "Vyberte soubor s daty pro elektro"
        //            }
        //            }
        //        }

        //    }
        //}

        public static string Místnost  {
            get {
                var Místnosti = Path.Combine(Informace.Instance.BasePath, "Místnosti");
                if (!Directory.Exists(Místnosti)) Directory.CreateDirectory(Místnosti);
                return Místnosti;
            }
        }
        public static string Revit {
            get {
                var Revit = Path.Combine(Místnost, "revit");
                if (!Directory.Exists(Revit)) Directory.CreateDirectory(Revit);
                return Revit;
            }
        }
        public static string MistnostiXLs => Path.Combine(Místnost, "Místnosti.celek.xlsx");
        public static string Mistnosti => Path.ChangeExtension(MistnostiXLs, ".json");

        /// <summary>
        /// Získá adresář se zdroji dat (stykače, měniče, kabely atd.).
        /// Pokud není nastaven v Informace.Instance.AdresarZdrojDat nebo neexistuje,
        /// použije se fallback na lokální složku ZdrojeDat v adresáři spuštění.
        /// </summary>
        public static string ZdrojDatAdresar
        {
            get
            {
                var cesta = Informace.Instance.AdresarZdrojDat;
                if (string.IsNullOrEmpty(cesta) || !Directory.Exists(cesta))
                {
                    cesta = Path.Combine(AdresarSpusteni ?? string.Empty, "ZdrojeDat");
                    if (!Directory.Exists(cesta))
                    {
                        Directory.CreateDirectory(cesta);
                    }
                }
                return cesta;
            }
        }

        //Cesta ke zdroji dat pro stykače, měniče a jističe, motory.
        public static string KM => Path.Combine(ZdrojDatAdresar, "Stykac", "KM.json");
        public static string KMCsv => Path.Combine(ZdrojDatAdresar, "Stykac", "KM.csv");

        public static string FM => Path.Combine(ZdrojDatAdresar, "Menic", "FM.json");
        public static string FMCsv => Path.Combine(ZdrojDatAdresar, "Menic", "FM.csv");

        public static string CuJson => Path.Combine(ZdrojDatAdresar, "Kabel", "Cu.json");
        public static string AlJson => Path.Combine(ZdrojDatAdresar, "Kabel", "Al.json");

        public static string JisticCsv => Path.Combine(ZdrojDatAdresar, "Jistic", "Jističe3VA.csv");
        public static string Jistic => Path.Combine(ZdrojDatAdresar, "Jistic", "Jističe3VA.json");

        public static string Motor => Path.Combine(ZdrojDatAdresar, "Motor", "MotoryList.json");
        public static string MotorCsv => Path.Combine(ZdrojDatAdresar, "Motor", "MotoryList.csv");
        public static string Motor3000Csv => Path.Combine(ZdrojDatAdresar, "Motor", "Motory3000.csv");

        public static string Motory => Path.Combine(ZdrojDatAdresar, "Motor", "Motory.Json");
        public static string MotoryCsv => Path.Combine(ZdrojDatAdresar, "Motor", "Motory.Csv");

        //Projekt
        public static string Projekt => Path.Combine(Informace.Instance.BasePath);
        public static string VyvodyOstatniJson => Path.Combine(Informace.Instance.BasePath, "Vyvody.Ostatni.json");
        public static string VyvodyTopeniJson => Path.Combine(Informace.Instance.BasePath, "Vyvody.Topeni.json");
        
        public static string VyvodyStavbaJson => Path.Combine(Informace.Instance.BasePath, "Vyvody.Stavba.json");
        
        public static string ElektroDataCsv => Path.Combine(Informace.Instance.BasePath, "ElektroData.Csv");
        
    }
}

