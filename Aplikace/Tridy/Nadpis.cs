using Aplikace.Tridy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Entiry
    {
        public int Id { get; set; }
        public string Apid { get; set; } = string.Empty;
    }

    public class Nadpis : Entiry
    {
        public string Name { get; set; }
        public string Jednotky { get; set; }

        public static List<Nadpis> dataEn() { return [
                new Nadpis { Name = "Equipment\nnumber", Jednotky=""  },
                new Nadpis { Name = "P&ID\nNumber", Jednotky="" },
                new Nadpis { Name = "Equipment name", Jednotky="" },
                new Nadpis { Name = "Power(electric)\n(EU Units)", Jednotky="[kW]" },
                new Nadpis { Name = "Package unit Power", Jednotky="" },
                new Nadpis { Name = "Variable speed drive", Jednotky="" },
                new Nadpis { Name = "PROUD Z TAB. PRO 500V", Jednotky="[A]" },
                new Nadpis { Name = "Power(electric)\n(US Units)", Jednotky="[HP]" },
                new Nadpis { Name = "CURRENT FOR 480V", Jednotky="[A]" },
                new Nadpis { Name = "COPPER CABLE SIZE\n(EU Units)", Jednotky="[mm2]" },
                new Nadpis { Name = "COPPER CABLE SIZE\n(US Units)", Jednotky="" },
                new Nadpis { Name = "CABLE LENGHT", Jednotky="[m]" },
                new Nadpis { Name = "DISTRIBUTOR EA/MCC", Jednotky="" },
                new Nadpis { Name = "DISTRIBUTOR NUMBER", Jednotky="" },
            ];
        }

        public static List<Nadpis> dataCz() { return [
                new Nadpis {Id=1, Name = "Označení", Jednotky=""  },
                new Nadpis {Id=2, Name = "Popis", Jednotky="" },
                new Nadpis {Id=3, Name = "Příkon", Jednotky="[kW]" },
                new Nadpis {Id=4, Name = "Proud", Jednotky="[A]" },
                new Nadpis {Id=5, Name = "Druh", Jednotky="" }, //BJ - balená jednotka, VSD- variabilní pohon, DOL- rozběh, Y/D
                new Nadpis {Id=6, Name = "Kabel", Jednotky="[mm2]" },
                new Nadpis {Id=7, Name = "Délka", Jednotky="[m]" },
                new Nadpis {Id=8, Name = "Rozvaděč", Jednotky="" },
                new Nadpis {Id=9, Name = "číslo", Jednotky="[m]" },
            ];
        }

    }
}
