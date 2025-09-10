using Aplikace.Sdilene;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Informace
    {
        public string BasePath { get; set; } = string.Empty;
        public string Místnost { get; set; } = string.Empty;
        public string Projekt { get; set; } = string.Empty;
        public string Název { get; set; } = string.Empty;
        public string Poznámka { get; set; } = string.Empty;
        public DateTime Datum { get; set; } = DateTime.Now;
    }

    public class InformaceProjektu
    {
        private static Informace informace;
        private InformaceProjektu() { }

        public static Informace Create() {
            
            //if (string.IsNullOrEmpty(informace.BasePath))
            //{ 
                //informace = new Informace();
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string file = Path.Combine(appData, "Elektro", "data.txt");
                if (File.Exists(file)) {
                    informace = Soubory.LoadJson<Informace>(file);
                }
                else {
                    informace = new Informace();
                }
            //}
            return informace;
        }
    }


}
