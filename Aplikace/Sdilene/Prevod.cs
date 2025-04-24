using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Reflection.PortableExecutable;
using System.Data;
using System.Text.Json.Nodes;
using Aplikace.Tridy;
using System.Reflection;

namespace Aplikace.Sdilene
{
    public static class Prevod
    {
        public static void DataTabletoToCsv(DataTable Table, string Soubor)
        {
            //DataTable Table = new() { TableName = "cestina", } ; 
            Table.TableName = "Cestina";
            using FileStream fs = new(Soubor, FileMode.Create);
            using StreamWriter sw = new(fs);
            string Pole = "";
            foreach (DataColumn item in Table.Columns)
            {
                Pole += item.ColumnName + ";";
            }
            sw.WriteLine(Pole[..^1]);

            foreach (DataRow item in Table.Rows)
            {
                Pole = "";
                foreach (DataColumn col in Table.Columns)
                {
                    Pole += item[col].ToString() + ";";
                }
                sw.WriteLine(Pole[..^1]);
            }
        }
        public static string JsonToCsv<T>(this List<T> json)
        {
            //string Json = JsonConvert.SerializeObject(json, Soubory.NastaveniEn());
            return JsonConvert.SerializeObject(json, Soubory.NastaveniEn());
            //JsonToCsv(Json, file);
        }
        public static void SaveToCsv<T>(this List<T> Class, string file)
        {
            string json = JsonToCsv(Class);
            SaveToCsv(json, file);
        }

        public static void SaveToCsv(string json, string file)
        {
            // Deserialize JSON to JArray
            JArray jsonArray = JArray.Parse(json);

            // Get property names from the first object (they will be used as headers)
            var headers = ((JObject)jsonArray[0]).Properties().Select(p => p.Name).ToArray();

            if(Soubory.IsFileLocked(file))
            {
                Console.WriteLine($"Soubor {file} je zamčený.");
                return;
            }

            // Prepare the CSV file
            //using (var writer = new StreamWriter(file, false ,Encoding.UTF8))
            using var writer = new StreamWriter(file, false, new UTF8Encoding(true));
            // Write the Hlavička
            writer.WriteLine(string.Join(";", headers));

            // Check if the array has elements
            if (jsonArray.Count > 0)
            {
                // Write data rows
                foreach (JObject obj in jsonArray.Cast<JObject>())
                {
                    //var values = obj.Properties().Select(p => p.Value.ToString()).ToArray();

                    //zachová entery \n
                    var values = obj.Properties()
                        .Select(p =>
                        {
                            var value = p.Value.ToString()
                                .Replace("\"", "\"\"")         // zdvoj uvozovky
                                .Replace("\n", " ")            // nebo zachovej \n, jak chceš
                                .Replace("\r", " ");           // odstran i \r, pokud je tam
                            return $"\"{value}\"";            // uzavři do uvozovek
                        })
                        .ToArray();
                    writer.WriteLine(string.Join(";", values));
                }
            }
            Console.WriteLine($"CSV soubor {Path.GetFileName(file)} byl vytvořen.");
        }


        //od umělé inteligence
        public static string JsonToXmlAI(string json)
        {
            // Zabalíme JSON, pokud začíná polem
            if (json.TrimStart().StartsWith('['))
                json = $"{{\"Hlavni\": {json}}}";

            // Použijeme přetíženou metodu, která specifikuje název kořenového elementu
            // Například "Root", můžete zvolit libovolný vhodný název
            //XDocument doc = JsonConvert.DeserializeXNode(json, "Polozka");

            // Parsuje JSON řetězec na objekt JObject
            JObject jObject = JsonConvert.DeserializeObject<JObject>(json);

            // Vytvoří se kořenový element XML dokumentu
            var xmlRoot = new XElement("Root");

            // Rekurzivně projde objekt JObject a přidá jeho prvky do XML
            AddJsonToXml(jObject, xmlRoot);

            // Vrátí se XML řetězec
            return xmlRoot.ToString();
        }

        /// <summary>Rekurzivně projde objekt JObject a přidá jeho prvky do XML</summary>
        private static void AddJsonToXml(JObject jObject, XElement parent)
        {
            foreach (var property in jObject.Properties())
            {
                var name = property.Name;
                var value = property.Value;

                var element = new XElement(name);

                if (value.Type == JTokenType.Object)
                {
                    // Pokud je hodnota objekt, rekurzivně se volá AddJsonToXml
                    AddJsonToXml((JObject)value, element);
                }
                else if (value.Type == JTokenType.Array)
                {
                    // Pokud je hodnota pole, rekurzivně se volá AddJsonToXml pro každý prvek v poli
                    foreach (var arrayValue in value.Children())
                    {
                        var arrayElement = new XElement("item");
                        AddJsonToXml((JObject)arrayValue, arrayElement);
                        element.Add(arrayElement);
                    }
                }
                else
                {
                    // Jinak se přidá hodnota jako textový element
                    element.Value = value.ToString();
                }

                parent.Add(element);
            }
        }

        /// <summary> Vstup Json jako string </summary>
        /// <returns>Výstup XML jako string</returns>
        public static string JsonToXml(string json)
        {
            // Zabalíme JSON, pokud začíná polem
            if (json.TrimStart().StartsWith('['))
            {
                json = $"{{\"Hlavni\": {json}}}";
            }

            // Použijeme přetíženou metodu, která specifikuje název kořenového elementu
            // Například "Root", můžete zvolit libovolný vhodný název
            XDocument doc = JsonConvert.DeserializeXNode(json, "Polozka");

            // Přidáme deklaraci XML, pokud není přítomna
            XDeclaration declaration = doc.Declaration ?? new XDeclaration("1.0", "utf-8", "yes");

            return $"{declaration}{Environment.NewLine}{doc}";
        }
 
        public static void UpdateCsvToJson<T>(List<T> sourceList, List<T> targets, string keyProperty = "Apid")
            where T : class, new()
        {
            var type = typeof(T);
            var keyProp = type.GetProperty(keyProperty) ?? throw new ArgumentException($"Property '{keyProperty}' not found.");
            // vytvoř slovník pro rychlé hledání dle klíče Apid
            var sourceDict = sourceList.ToDictionary(
                item => keyProp.GetValue(item)?.ToString() ?? string.Empty
            );

            // všechny veřejné zapisovatelné vlastnosti kromě klíče a Item
            //item nevím co to je
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => p.CanWrite && p.Name != keyProperty && p.Name != "Item")
                                 .ToList();

            foreach (var target in targets)
            {
                var key = keyProp.GetValue(target)?.ToString() ?? string.Empty;

                if (sourceDict.TryGetValue(key, out var source))
                {
                    foreach (var prop in properties)
                    {
                        var value = prop.GetValue(source);
                        prop.SetValue(target, value);
                    }
                }
            }
        }

        public static void Vypis(this List<List<string>> Pole)
        {
            foreach (var item in Pole)
            {
                foreach (var i in item)
                {
                    Console.Write($"{i}, ");
                }
                Console.Write("\n");
            }
        }
    }
}
