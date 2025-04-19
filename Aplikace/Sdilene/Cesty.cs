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
    }
}
