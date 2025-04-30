using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    //Data jsou ze stránky ABB
    //https://search.abb.com/library/Download.aspx?DocumentID=CZZPA200516021233-J&LanguageCode=cs&DocumentPartId=1&Action=Launch
    public class Menic
    {
        public double Prikon { get; set; }                     // Příkon ve kW
        public double PrikonHP { get; set; }                   // Příkon v HP
        public double Proud { get; set; }                      // Proud v A
        public string Provoz { get; set; } = string.Empty;     // Režim provozu (Normální nebo Těžký)
        public string TypovyKod { get; set; } = string.Empty;  // Typový kód měniče
        public string Velikost { get; set; } = string.Empty;   // Velikost rámu (např. R1, R2, ...)
        public int NapetiMin { get; set; }                     // Minimální napětí
        public int NapetiMax { get; set; }                     // Maximální napětí
        public string Výrobce { get; set; } = string.Empty;

        public static List<Menic> Nacti(string cesta)
        {
            var vysledek = new List<Menic>();
            var lines = File.ReadAllLines(cesta);
            var culture = new CultureInfo("cs-CZ");

            foreach (var line in lines.Skip(1)) // přeskočíme hlavičku
            {
                var parts = line.Split(';');
                vysledek.Add(new Menic
                {
                    Prikon = double.Parse(parts[0], culture),
                    PrikonHP = double.Parse(parts[1], culture),
                    Proud = double.Parse(parts[2], culture),
                    Provoz = parts[3],
                    TypovyKod = parts[4],
                    Velikost = parts[5],
                    NapetiMin = int.Parse(parts[6]),
                    NapetiMax = int.Parse(parts[7])
                });
            }

            return vysledek;
        }

        public static void Uloz(string cesta, List<Menic> menice)
        {
            var culture = new CultureInfo("cs-CZ");
            var sb = new StringBuilder();

            // Hlavička
            sb.AppendLine("Příkon;PříkonHP;Proud;Provoz;Typový kód;Velikost;Napětí;Napětí");

            foreach (var m in menice)
            {
                sb.AppendLine(string.Join(";", new string[]
                {
                m.Prikon.ToString(culture),
                m.PrikonHP.ToString(culture),
                m.Proud.ToString(culture),
                m.Provoz,
                m.TypovyKod,
                m.Velikost,
                m.NapetiMin.ToString(),
                m.NapetiMax.ToString()
                }));
            }

            File.WriteAllText(cesta, sb.ToString(), Encoding.UTF8);
        }

    }
}
