using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Aplikace.Sdilene
{
    public static class Soubory
    {
        readonly static Encoding čeština = Encoding.GetEncoding(1250); //funguje čeština

        public static JsonSerializerSettings Nastaveni()
        {
            var settings = new JsonSerializerSettings()
            {
                Culture = System.Globalization.CultureInfo.GetCultureInfo("cs-CZ"),
                //formátovaný text
                Formatting = Newtonsoft.Json.Formatting.Indented,
                //ignorovány vlastnosti s hodnotou null.
                NullValueHandling = NullValueHandling.Ignore,
                //FloatParseHandling = FloatParseHandling.Decimal,
                //ignorováno kdy JSON není v cílovém typu objektu.
                MissingMemberHandling = MissingMemberHandling.Ignore,

                ReferenceLoopHandling = ReferenceLoopHandling.Ignore,

                //DefaultValueHandling = DefaultValueHandling.Ignore,
                //ContractResolver = new IgnoreEmptyStringResolver(),
            };
            return settings;
        }

        public static JsonSerializerSettings NastaveniEn()
        {
            var settings = new JsonSerializerSettings()
            {
                //Culture = System.Globalization.CultureInfo.GetCultureInfo("cs-CZ"),
                //formátovaný text
                Formatting = Newtonsoft.Json.Formatting.Indented,

                //ignorovány vlastnosti s hodnotou null.
                NullValueHandling = NullValueHandling.Ignore,

                //FloatParseHandling = FloatParseHandling.Decimal,

                //ignorováno kdy JSON není v cílovém typu objektu.
                //MissingMemberHandling = MissingMemberHandling.Ignore,

                //Určuje možnosti zpracování referenční smyčky pro JsonSerializer.
                //Ignorujte odkazy na smyčky a neprovádějte serializaci.
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore,

                //DefaultValueHandling = DefaultValueHandling.Ignore,

                //ContractResolver = new IgnoreEmptyStringResolver(),
            };
            return settings;
        }

        /// <summary>
        /// uložit soubor, deserializace třídy pozor na vstup generika
        /// </summary>
        public static void SaveJsonList<T>(this List<T> values, string cesta) where T : class
        {
            //MessageBox.Show("save");
            // Nastavení formátování JSON s odsazením (entery)
            //var settings = new JsonSerializerSettings { Formatting = Formatting.Indented };
            string Json = JsonConvert.SerializeObject(values, Nastaveni());
            //MessageBox.Show(Json);
            File.WriteAllText(cesta, Json);
            return;
        }

        /// <summary> načti soubor uvedená třida doplněna do LIST , deserializace třídy pozor na vstup generika  </summary>
        public static List<T> LoadJsonList<T>(string cesta) where T : class
        {
            if (System.IO.File.Exists(cesta))
            {
                string jsonString = System.IO.File.ReadAllText(cesta);
                List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString, Nastaveni()) ?? [];
                return moje;
            }
            return [];
        }

        /// <summary> načti soubor uvedená třida doplněna do LIST , deserializace třídy pozor na vstup generika  </summary>
        public static T LoadJson<T>(string cesta) where T : new() 
        {
            if (System.IO.File.Exists(cesta))
            {
                string jsonString = System.IO.File.ReadAllText(cesta);
                T moje = Newtonsoft.Json.JsonConvert.DeserializeObject<T>(jsonString, Nastaveni()) ?? new();
                return moje;
            }
            return new();
        }

        /// <summary> načti soubor uvedená třida doplněna do LIST , deserializace třídy pozor na vstup generika </summary>
        public static List<T> LoadJsonEn<T>(string cesta) where T : new()
        {
            if (System.IO.File.Exists(cesta))
            {
                string jsonString = System.IO.File.ReadAllText(cesta);
                //List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString);
                List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString, NastaveniEn()) ?? [];
                return moje;
            }
            return [];
        }

        public static void KillExcel()
        {
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                try
                {
                    process.Kill();
                    process.WaitForExit();
                    Console.WriteLine($"Proces {process.Id} ukončen");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chyba při ukončování: {ex.Message}");
                }
            }
        }
        public static void KillExcel(int processId)
        {
            var process = Process.GetProcessById(processId);
                try
                {
                    process.Kill();
                    process.WaitForExit();
                    Console.WriteLine($"Proces {process.Id} ukončen");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chyba při ukončování: {ex.Message}");
                }

        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        //[LibraryImport("user32.dll", SetLastError = true)]
        //public static partial uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static int GetExcelProcess(Application myApp)
        {
            //IntPtr myHwnd = (IntPtr)myApp.Hwnd;
            IntPtr myHwnd = (IntPtr)myApp.Hwnd;
            GetWindowThreadProcessId(myHwnd, out uint myProcessId);
            return (int)myProcessId;
        }

        public static bool IsFileLocked(string path)
        {
            try
            {
                using FileStream stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                // Pokud se otevře, není zamčený
                return false;
            }
            catch (IOException)
            {
                // Pokud dojde k výjimce, soubor je pravděpodobně zamčený jiným procesem
                return true;
            }
        }
    }
}
