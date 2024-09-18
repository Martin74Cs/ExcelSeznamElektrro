using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Aplikace.Sdilene
{
    public static class Soubory
    {
        public static JsonSerializerSettings nastaveni()
        {
            JsonSerializerSettings settings = new JsonSerializerSettings()
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

        /// <summary>
        /// uložit soubor, deserializace třídy pozor na vstup generika
        /// </summary>
        public static void SaveJsonList<T>(this List<T> values, string cesta) where T : class
        {
            //MessageBox.Show("save");
            // Nastavení formátování JSON s odsazením (entery)
            //var settings = new JsonSerializerSettings { Formatting = Formatting.Indented };
            string Json = JsonConvert.SerializeObject(values, nastaveni());
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
                List<T> moje = Newtonsoft.Json.JsonConvert.DeserializeObject<List<T>>(jsonString, nastaveni());
                return moje;
            }
            return new List<T>();
        }
    }
}
