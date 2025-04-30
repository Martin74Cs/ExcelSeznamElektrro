using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Aplikace.Sdilene
{
    public static partial class Soubory
    {
        //readonly static Encoding čeština = Encoding.GetEncoding(1250); //funguje čeština

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
                StringEscapeHandling = StringEscapeHandling.Default,

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
            Console.WriteLine($"Json soubor {Path.GetFileName(cesta)} byl vytvořen.");
            return;
        }

        /// <summary> Načti soubor uvedená třida doplněna do LIST , deserializace třídy pozor na vstup generika  </summary>
        public static List<T> LoadJsonList<T>(string cesta) where T : class
        {
            if (!System.IO.File.Exists(cesta)) return [];
            string jsonString = System.IO.File.ReadAllText(cesta);
            List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString, Nastaveni()) ?? [];
            return moje;
        }

        public static List<T> LoadJsonListEn<T>(string cesta) where T : class
        {
            if (!System.IO.File.Exists(cesta)) return [];
            string jsonString = System.IO.File.ReadAllText(cesta);
            List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString, NastaveniEn()) ?? [];
            return moje;

        }

        /// <summary> Načti soubor uvedená třida doplněna do LIST , deserializace třídy pozor na vstup generika. </summary>
        public static T LoadJson<T>(string cesta) where T : new() 
        {
            if (!System.IO.File.Exists(cesta)) return new();
            string jsonString = System.IO.File.ReadAllText(cesta);
            T moje = Newtonsoft.Json.JsonConvert.DeserializeObject<T>(jsonString, Nastaveni()) ?? new();
            return moje;
        }

        /// <summary> načti soubor uvedená třida doplněna do LIST , deserializace třídy pozor na vstup generika </summary>
        public static List<T> LoadJsonEn<T>(string cesta) where T : new()
        {
            if (!System.IO.File.Exists(cesta)) return [];
            string jsonString = System.IO.File.ReadAllText(cesta);
            //List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString);
            List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString, NastaveniEn()) ?? [];
            return moje;
        }


        public static List<T> LoadFromCsv<T>(string file ) where T : new()
        {
            if (!File.Exists(file)) return [];
            var list = new List<T>();

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            //using (var reader = new StreamReader(file, Encoding.UTF8))
            using (var reader = new StreamReader(file, Encoding.GetEncoding(1250)))
            {
                //Načti hlavičku
                var headerLine = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(headerLine))
                    return list; // prázdný soubor

                var headers = headerLine.Split(';');
                var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

                //Načti data
                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    var values = ParseCsvLine(line, headers.Length);
                    var obj = new T();
                    for (int i = 0; i < headers.Length && i < values.Count; i++)
                    {
                        var header = headers[i];
                        //Najdi v seznamu vlastností první vlastnost, jejíž název(p.Name) odpovídá názvu sloupce(header) z CSV, bez ohledu na velikost písmen(case -insensitive).
                        var property = properties.FirstOrDefault(p => 
                            string.Equals(p.Name, header, StringComparison.OrdinalIgnoreCase));
                        if(header == "Obvod" || header == "Plocha" || header == "Objem") 
                            values[i] = values[i].Split(' ').FirstOrDefault(); // odstranění uvozovek
                        if (string.IsNullOrEmpty(values[i])) continue;

                        if (property != null && property.CanWrite)
                        {
                            try
                            {
                                object? convertedValue = Convert.ChangeType(values[i], property.PropertyType);
                                property.SetValue(obj, convertedValue);
                            }
                            catch { // Pokud převod selže, můžeš logovat nebo nastavit výchozí hodnotu
                                    }
                        }
                    }
                     list.Add(obj);
                }
            }
            return list; // nebo jsonArray.ToString(Formatting.Indented) pro čitelný výstup
        }

        private static List<string> ParseCsvLine(string line, int expectedColumns)
        {
            var values = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"' && (i == 0 || line[i - 1] != '\\'))
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"'); // zdvojené uvozovky = 1 uvozovka
                        i++;
                    }
                    else
                        inQuotes = !inQuotes;
                }
                else if (c == ';' && !inQuotes)
                {
                    values.Add(sb.ToString());
                    sb.Clear();
                }
                else
                    sb.Append(c);
            }

            values.Add(sb.ToString());

            // Doplnění prázdných sloupců, pokud jich je méně než hlaviček
            while (values.Count < expectedColumns)
                values.Add("");

            return values;
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

        [LibraryImport("user32.dll", SetLastError = true)]
        //[LibraryImport("user32.dll", SetLastError = true)]
        private static partial uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

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
            if (!File.Exists(path)) return false;
            try {
                using FileStream stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                // Pokud se otevře, není zamčený
                return false;
            }
            catch (IOException) {
                // Pokud dojde k výjimce, soubor je pravděpodobně zamčený jiným procesem
                return true;
            }
        }
    }
}
