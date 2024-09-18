using System;
using System.Collections.Generic;
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
    }
}
