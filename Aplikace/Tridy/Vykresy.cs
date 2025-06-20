using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

using System.Threading.Tasks;

namespace Aplikace.Tridy {
    public class Vykres {
        public string Orientačníčíslo { get {
                if(string.IsNullOrEmpty(OrientačníčísloF))
                    return $"{OrientačníčísloB}{OrientačníčísloC}{OrientačníčísloD}{OrientačníčísloE}";
                return $"{OrientačníčísloB}{OrientačníčísloC}{OrientačníčísloD}{OrientačníčísloE}-{OrientačníčísloF}";
        } } 

        [JsonIgnore]
        public string OrientačníčísloB { get; init; }  = string.Empty;
        [JsonIgnore]
        public string OrientačníčísloC { get; init; } = string.Empty;
        [JsonIgnore]
        public string OrientačníčísloD { get; init; } = string.Empty;
        [JsonIgnore]
        public string OrientačníčísloE { get; init; } = string.Empty;
        [JsonIgnore]
        public string OrientačníčísloF { get; init; } = string.Empty;

        public string ČísloDokumentu { get; set; } = string.Empty;
        public string Nazev { get; set; } = string.Empty;
        public string ProfesníČislo { get; set; } = string.Empty;
        public string Revize { get; set; } = string.Empty;
        public string Popisrevize { get; set; } = string.Empty;
        public string Cesta { get; set; } = string.Empty;

        /// <summary>Volání parametru jako string např. Nadpis[Name]  </summary>
        [JsonIgnore]
        public object? this[string nazev]
        {
            get
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance) ?? throw new ArgumentException($"Neexistující vlastnost: {nazev}");
                return prop.GetValue(this);
            }
            set
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance) ?? throw new ArgumentException($"Neexistující vlastnost: {nazev}");
                prop.SetValue(this, Convert.ChangeType(value, prop.PropertyType));
            }
        }
    }
}
