using Aplikace.Tridy;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Entity
    {
        public int Id { get; set; }
        public string Apid { get; set; } = string.Empty;
    }

    public class Nadpis : Entity
    {
        [Display(Name = "Text Nadpisu")]
        public string Name { get; set; } = string.Empty;

        [Display(Name = "Jednotky")]
        public string Jednotky { get; set; } = string.Empty;

        public static List<Nadpis> DataEn() { return [
                new Nadpis {Id=1,  Name = "Equipment\nnumber",              Jednotky=""  },
                new Nadpis {Id=2,  Name = "P&ID\nNumber",                   Jednotky="" },
                new Nadpis {Id=3,  Name = "Equipment name",                 Jednotky="" },
                new Nadpis {Id=4,  Name = "Power(electric)\n(EU Units)",    Jednotky="[kW]" },
                new Nadpis {Id=5,  Name = "Package unit Power",             Jednotky="" },
                new Nadpis {Id=6,  Name = "Variable speed drive",           Jednotky="" },
                new Nadpis {Id=7,  Name = "PROUD Z TAB. PRO 500V",          Jednotky="[A]" },
                new Nadpis {Id=8,  Name = "Power(electric)\n(US Units)",    Jednotky="[HP]" },
                new Nadpis {Id=9,  Name = "CURRENT FOR 480V",               Jednotky="[A]" },
                new Nadpis {Id=10, Name = "COPPER CABLE SIZE\n(EU Units)",  Jednotky="[mm2]" },
                new Nadpis {Id=11, Name = "COPPER CABLE SIZE\n(US Units)",  Jednotky="" },
                new Nadpis {Id=12, Name = "CABLE LENGHT",                   Jednotky="[m]" },
                new Nadpis {Id=13, Name = "DISTRIBUTOR EA/MCC",             Jednotky="" },
                new Nadpis {Id=14, Name = "DISTRIBUTOR NUMBER",             Jednotky="" },
            ];
        }

        public static List<Nadpis> DataCz() { return [
                new Nadpis {Id=1, Name = "Označení",    Jednotky=""  },
                new Nadpis {Id=2, Name = "Popis",       Jednotky="" },
                new Nadpis {Id=3, Name = "Příkon",      Jednotky="[kW]" },
                new Nadpis {Id=4, Name = "Napětí",      Jednotky="[V]" },
                new Nadpis {Id=5, Name = "Proud",       Jednotky="[A]" },
                new Nadpis {Id=6, Name = "Balená",       Jednotky="" },
                new Nadpis {Id=7, Name = "Měnič",       Jednotky="" },
                new Nadpis {Id=8, Name = "Druh",        Jednotky="" }, //BJ - balená jednotka, VSD- variabilní pohon, DOL- rozběh, Y/D
                new Nadpis {Id=9, Name = "Kabel",       Jednotky="[mm2]" },
                new Nadpis {Id=10, Name = "Délka",       Jednotky="[m]" },
                new Nadpis {Id=11, Name = "Rozvaděč",    Jednotky="" },
                new Nadpis {Id=12, Name = "číslo",       Jednotky="" },
            ];
        }
        /// <summary>Volání parametru jako string např. Nadpis[Name]  </summary>
        public object this[string nazev]
        {
            get
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance);
                return prop == null ? throw new ArgumentException($"Neexistující vlastnost: {nazev}") : prop.GetValue(this);
            }
            set
            {
                var prop = GetType().GetProperty(nazev, BindingFlags.Public | BindingFlags.Instance) ?? throw new ArgumentException($"Neexistující vlastnost: {nazev}");
                prop.SetValue(this, Convert.ChangeType(value, prop.PropertyType));
            }
        }

        /// <summary> List vlastností třídy </summary>
        public List<string> Vlastnosti => GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                      .Select(p => p.Name)
                      .ToList();

    }
}
