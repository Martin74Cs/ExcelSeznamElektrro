using Aplikace.Sdilene;
using Aplikace.Tridy;
using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Formats.Asn1;
using System.Globalization;
using System.IO;
using System.Linq;
using static Oez;

public class Oez
{
    public static List<Product> LoadProductsFromCsv(string filePath)
    {
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = ";",
            Encoding = System.Text.Encoding.UTF8,
            BadDataFound = null,
            MissingFieldFound = null
        };

        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, config))
        {
            csv.Context.RegisterClassMap<ProductMap>();
            var products = csv.GetRecords<Product>().ToList();

            // Parsování Popis
            foreach (var product in products)
            {
                ParseProductDescription(product);
            }

            return products;
        }
    }

    private static void ParseProductDescription(Product product)
    {
        if (string.IsNullOrWhiteSpace(product.Popis))
            return;

        // Rozdělíme popis na části podle čárky
        var parts = product.Popis.Split(", ")
            .Select(p => p.Trim())
            .ToList();

        foreach (var part in parts)
        {
            // Proud (In X A)
            if (part.StartsWith("In ") && part.EndsWith(" A"))
            {
                if (double.TryParse(part.Replace("In ", "").Replace(" A", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var current))
                    product.InA = current;
            }
            // Proud (Ie X A)
            else if (part.StartsWith("Ie ") && part.EndsWith(" A"))
            {
                if (double.TryParse(part.Replace("Ie ", "").Replace(" A", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var current))
                    product.InA = current;
            }
            // Napětí (Ue ...)
            else if (part.StartsWith("Ue "))
            {
                product.Ue = part.Replace("Ue ", "").Replace(" V", "");
            }
            // Charakteristika
            else if (part.StartsWith("charakteristika "))
            {
                product.Characteristic = part.Replace("charakteristika ", "");
            }
            // Zbytkový proud (Idn X mA)
            else if (part.StartsWith("Idn ") && part.EndsWith(" mA"))
            {
                if (double.TryParse(part.Replace("Idn ", "").Replace(" mA", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var idn))
                    product.Idn = idn;
            }
            // Počet pólů
            else if (part.Contains("pól"))
            {
                product.Poly = part;
            }
            // Šířka v modulech
            else if (part.StartsWith("šířka ") && part.EndsWith(" modul"))
            {
                if (double.TryParse(part.Replace("šířka ", "").Replace(" modul", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var width))
                    product.Moduly = width;
            }
            // Zkratová schopnost (Icn X kA)
            else if (part.StartsWith("Icn ") && part.EndsWith(" kA"))
            {
                if (double.TryParse(part.Replace("Icn ", "").Replace(" kA", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var icn))
                    product.Icn = icn;
            }
            else if (part.StartsWith("pro válcové pojistkové "))
            {
                    product.Vlozka = part.Replace("pro válcové pojistkové vložky ", "");
            }
            // Typ (např. typ A, typ AC)
            else if (part.StartsWith("typ "))
            {
                product.Type = part.Replace("typ ", "");
            }
            // Typ (např. typ A, typ AC)
            else if (part.Contains("signaliz"))
            {
                product.Signalizace = part;
            }

        }
    }

// Mapování CSV sloupců
    public class ProductMap : ClassMap<Product>
    {
        public ProductMap()
        {
            Map(m => m.Kod).Name("Objednací kód");
            Map(m => m.Skupina).Name("Typové značení");
            Map(m => m.Jmeno).Name("Název produktu");
            Map(m => m.Popis).Name("Popis produktu");
            //Map(m => m.Popis).Name("Popis produktu");
            Map(m => m.Rabat).Name("Rabatová skupina");
            Map(m => m.ZakladniCena).Name("Základní cena");
            Map(m => m.Jednotka).Name("Měrná jednotka položky");
            Map(m => m.EAN).Name("EAN");
            Map(m => m.Objem).Name("Objem (brutto) [dm3]");
            Map(m => m.Hmotnost).Name("Hmotnost (brutto) [kg]");
            Map(m => m.Vyska).Name("Výška (brutto) [mm]");
            Map(m => m.Sirka).Name("Šířka (brutto) [mm]");
            Map(m => m.Delka).Name("Délka (brutto) [mm]");
            Map(m => m.CustomsCode).Name("Celní kód (intrastat)");
            Map(m => m.ZemePuvodu).Name("Země původu");
            Map(m => m.AL).Name("AL");
            Map(m => m.ECCN).Name("ECCN");
            Map(m => m.KusyBaleni).Name("Počet kusů v základním balení");
            Map(m => m.KusyMin).Name("Nedělitelné Objednací množství");
        }
    }
    // Mapování CSV sloupců
    public class Product
    {
        // Základní vlastnosti z CSV
        public string Kod { get; set; }
        public string Skupina { get; set; }
        public string Jmeno { get; set; }
        public string Popis { get; set; } // Původní popis pro referenci
        public string Rabat { get; set; }
        public decimal ZakladniCena { get; set; }
        public string Jednotka { get; set; }
        public string EAN { get; set; }
        public double Objem { get; set; }
        public double Hmotnost { get; set; }
        public double Vyska { get; set; }
        public double Sirka { get; set; }
        public double Delka { get; set; }
        public string CustomsCode { get; set; }
        public string ZemePuvodu { get; set; }
        public string AL { get; set; }
        public string ECCN { get; set; }
        public double? KusyBaleni { get; set; }
        public double KusyMin { get; set; }

        // Rozparsované atributy z Popis
        public double? InA { get; set; } // In (proud, např. 10 A)
        public string Ue { get; set; } // Ue (napětí, např. AC 230 V)
        public string Characteristic { get; set; } // Charakteristika (např. B, C)
        public double? Idn { get; set; } // Idn (zbytkový proud, např. 30 mA)
        public string Poly { get; set; } // Počet pólů (např. 1+N-pól, 1pól)
        public double? Moduly { get; set; } // Šířka v modulech
        public double? Icn { get; set; } // Jmenovitá zkratová schopnost (např. 6 kA)
        public string Type { get; set; } // Typ (např. A, AC, A-G)
        public string Vlozka { get; set; } // pro válcové pojistkové vložky 14x51
        public string Signalizace { get; set; } // pro válcové pojistkové vložky 14x51
    }

}

public class Priklad
{
     public static List<Product> FindCircuitBreakers(List<Product> products, 
        double? currentInA = null, 
        string voltageUe = null, 
        string characteristic = null, 
        string type = null) {
        return products
            .Where(p => p.Jmeno.Contains("Jistič") || p.Jmeno.Contains("Jističochránič"))
            .Where(p => currentInA == null || p.InA == currentInA)
            .Where(p => voltageUe == null || p.Ue == voltageUe)
            .Where(p => characteristic == null || p.Characteristic == characteristic)
            .Where(p => type == null || p.Type == type)
            .ToList();
    }

    public static void Main()
    {
        string file = "OEZExportZbozi2025-05-29.csv";
        string Cesta = Path.Combine(Cesty.Data, "Jištení" , file);
        List<Product> products = LoadProductsFromCsv(Cesta);

        var Pole = products.GroupBy(x => x.Skupina).OrderBy(x => x.Key);
        foreach(var item in Pole) 
            //Console.WriteLine(item.Key);

        Console.WriteLine("\nDalší");
        var Pojistka = Pole.Where(x => x.Key.Contains("Pojistková vložka")).SelectMany(g => g).ToList();
        foreach (var item in Pojistka)
            //Console.WriteLine(item.Jmeno);

        Console.WriteLine("\nDalší");
        var PojOdpínač = Pole.Where(x => x.Key.Contains("Pojistkový odpínač")).SelectMany(g => g).ToList();
        foreach (var item in PojOdpínač)
            Console.WriteLine(item.Jmeno);

        foreach (var item in PojOdpínač) { 
            Console.WriteLine(item);
            VypisVlastnosti(item);
        }


        // Najít jističe s proudem 16 A, napětím AC 230 V a charakteristikou B
        var breakers = FindCircuitBreakers(products, currentInA: 16, voltageUe: "AC 230", characteristic: "B");
        Console.WriteLine($"Nalezeno {breakers.Count} jističů (16 A, AC 230 V, charakteristika B):");
        foreach (var breaker in breakers)
        {
            Console.WriteLine($"Kód: {breaker.Kod}, Název: {breaker.Jmeno}, Cena: {breaker.ZakladniCena}");
        }

        // Najít jističe typu A-G
        var agBreakers = FindCircuitBreakers(products, type: "A-G");
        Console.WriteLine($"\nNalezeno {agBreakers.Count} jističů typu A-G:");
        foreach (var breaker in agBreakers)
        {
            Console.WriteLine($"Kód: {breaker.Kod}, Popis: {breaker.Popis}");
        }

    }

    public static void VypisVlastnosti<T>(T obj)
    {
        if (obj == null)
        {
            Console.WriteLine("Objekt je null.");
            return;
        }

        var typ = typeof(T);
        var vlastnosti = typ.GetProperties();

        Console.WriteLine($"Výpis vlastností objektu typu {typ.Name}:");

        foreach (var vlastnost in vlastnosti)
        {
            var hodnota = vlastnost.GetValue(obj);
            Console.WriteLine($"{vlastnost.Name,-15} = {hodnota}");
        }
    }
}