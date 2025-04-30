using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

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
        public static string Cesta { get =>  Path.Combine(UserProfile , @"AppData\Roaming\Autodesk\ApplicationPlugins\Elektro.bundle"); }
        public static string CestaSich => Path.Combine(Cesta , "Sichr");
        public static string CuJson => Path.Combine(CestaSich , "Cu.Json");
        public static string AlJson => CestaSich + @"Al.Json";

        /// <summary>...LightChem\Elektro\Lightchem </summary>
        public static string Lightchem => Path.Combine(BasePath , "Lightchem");

        /// <summary>...Můj disk\Elektro\Lightchem\Ëlektro </summary>
        public static string Elektro => Path.Combine(Lightchem , "Elektro");
        public static string Data => Path.Combine(Lightchem, "Data");
        public static string MotoryJson => Path.Combine(Data, "Motory", "Motory.Json");
        public static string ElektroDataCsv => Path.Combine(Elektro, "ElektroData.Csv");
        public static string ElektroDataJson => Path.Combine(Elektro, "ElektroData.Json");
        public static string ElektroRozvaděčJson => Path.Combine(Elektro, "ElektroRozvaděč.Json");
        
        public static string BasePath {
            get {
                if (Environment.UserDomainName == "D10")
                    return @"c:\a\";
                else
                    return @"G:\Můj disk\Projekty\";
            }
        }
        public static string GooglePath {
            get {
                if (Environment.UserDomainName == "D10")
                    return @"E:\Můj disk\Projekty\";
                else
                    return @"G:\Můj disk\Projekty\";
            }
        }

        public static string Místnost  {
            get {
                var Místnosti = Path.Combine(Lightchem, "Místnosti");
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
        public static string MistnostiJson => Path.ChangeExtension(MistnostiXLs, ".json");
    }
}
