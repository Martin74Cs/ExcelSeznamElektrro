using Aplikace.Tridy;
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
        public static string Cesta { get => Path.Combine(UserProfile, @"AppData\Roaming\Autodesk\ApplicationPlugins\Elektro.bundle"); }
        public static string CestaSichAcad => Path.Combine(Cesta, "Sichr");
        public static string CuJsonAcad => Path.Combine(CestaSichAcad, "Cu.Json");
        public static string AlJsonAcad => CestaSichAcad + @"Al.Json";

        /// <summary>...LightChem\Elektro\Lightchem </summary>
        //public static string Lightchem => Path.Combine(BasePath, "Lightchem");

        /// <summary>...Můj disk\Elektro\Lightchem\Ëlektro </summary>
        //public static string Elektro { get {
        //        //if (!Directory.Exists(Cesty.Elektro))
        //        //    Directory.CreateDirectory(Cesty.Elektro);
        //        return Path.Combine(Lightchem, "Elektro");
        //    }
        //}

        //public static string ElektroDataCsv => Path.Combine(Elektro, "ElektroData.Csv");
        //public static string ElektroDataJson => Path.Combine(Elektro, "ElektroData.Json");
        //public static string ElektroRozvaděčJson => Path.Combine(Elektro, "ElektroRozvaděč.Json");
        
        //public static string BasePath {
        //    get {
        //        //if (Environment.UserDomainName == "D10")
        //        //    //return @":\a\";
        //        //    return @"E:\Můj disk\Projekty\";
        //        //else
        //        //    return @"G:\Můj disk\Projekty\";

        //        using var Inforamce = Informace.Create;
        //        if (!Directory.Exists(Inforamce.AdresarZdrojDat)) {
        //            OpenFileDialog dialog = new() {
        //                Title = "Vyberte soubor s daty pro elektro"
        //                //Filter = "Json files (*.json)|*.json|All files (*.*)|*.*";
        //            };
        //            if (dialog.ShowDialog() == DialogResult.OK) {
        //                Inforamce.AdresarZdrojDat = Path.GetDirectoryName(dialog.FileName) ?? string.Empty;
        //                //MessageBox.Show($"Vybrali jste soubor: {dialog.FileName}");
        //            }
        //            else {
        //                MessageBox.Show("Nebyl vybrán žádný soubor. Aplikace bude ukončena.");
        //                Environment.Exit(0);
        //            }
        //        }

        //        return Inforamce.AdresarZdrojDat;
        //    }
        //}

        public static string Místnost  {
            get {
                var Místnosti = Path.Combine(Informace.Create.BasePath, "Místnosti");
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
        //public static string Data => Path.Combine(BasePath, "Data");
        public static string MistnostiXLs => Path.Combine(Místnost, "Místnosti.celek.xlsx");
        public static string MistnostiJson => Path.ChangeExtension(MistnostiXLs, ".json");

        //Cesta ke zdroji dat pro stykače, měniče a jističe, motory.
        public static string CestaKM => Path.Combine(Informace.Create.AdresarZdrojDat,"Stykac", "KM.csv");

        public static string CestaFM => Path.Combine(Informace.Create.AdresarZdrojDat, "Menic", "FM.csv");

        public static string CuJson => Path.Combine(Informace.Create.AdresarZdrojDat, "Kabel", "Cu.json");
        public static string AlJson => Path.Combine(Informace.Create.AdresarZdrojDat, "Kabel", "Al.json");

        public static string CestaJistic => Path.Combine(Informace.Create.AdresarZdrojDat, "Jistic", "Jističe3VA.csv");

        public static string CestaMotor => Path.Combine(Informace.Create.AdresarZdrojDat, "Motor", "MotoryList.json");
        public static string CestaMotorCsv => Path.Combine(Informace.Create.AdresarZdrojDat, "Motor", "MotoryList.csv");
        public static string CestaMotor3000Csv => Path.Combine(Informace.Create.AdresarZdrojDat, "Motor", "Motory3000.csv");


        //public static string MotoryJson => Path.Combine(Informace.Create.AdresarZdrojDat, "Motory", "Motory.Json");
        

        //Projekt
        public static string Projekt => Path.Combine(Informace.Create.BasePath);
        public static string VyvodyJson => Path.Combine(Informace.Create.BasePath, "Vývody.json");

        public static string VyvodyStavbaJson => Path.Combine(Informace.Create.BasePath, "Vývody.Stavba.json");
        
        public static string ElektroDataCsv => Path.Combine(Informace.Create.BasePath, "ElektroData.Csv");
        
    }
}
