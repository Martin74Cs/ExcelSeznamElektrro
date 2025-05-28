using Aplikace.Sdilene;
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

            // Parsování ProductDescription
            foreach (var product in products)
            {
                ParseProductDescription(product);
            }

            return products;
        }
    }

    private static void ParseProductDescription(Product product)
    {
        if (string.IsNullOrWhiteSpace(product.ProductDescription))
            return;

        // Rozdělíme popis na části podle čárky
        var parts = product.ProductDescription.Split(", ")
            .Select(p => p.Trim())
            .ToList();

        foreach (var part in parts)
        {
            // Proud (In X A)
            if (part.StartsWith("In ") && part.EndsWith(" A"))
            {
                if (double.TryParse(part.Replace("In ", "").Replace(" A", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var current))
                    product.CurrentInA = current;
            }
            // Napětí (Ue ...)
            else if (part.StartsWith("Ue "))
            {
                product.VoltageUe = part.Replace("Ue ", "").Replace(" V", "");
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
                    product.ResidualCurrentIdn = idn;
            }
            // Počet pólů
            else if (part.Contains("-pól"))
            {
                product.Poles = part;
            }
            // Šířka v modulech
            else if (part.StartsWith("šířka ") && part.EndsWith(" modul"))
            {
                if (double.TryParse(part.Replace("šířka ", "").Replace(" modul", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var width))
                    product.WidthInModules = width;
            }
            // Zkratová schopnost (Icn X kA)
            else if (part.StartsWith("Icn ") && part.EndsWith(" kA"))
            {
                if (double.TryParse(part.Replace("Icn ", "").Replace(" kA", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var icn))
                    product.BreakingCapacityIcn = icn;
            }
            // Typ (např. typ A, typ AC)
            else if (part.StartsWith("typ "))
            {
                product.Type = part.Replace("typ ", "");
            }
        }
    }

// Mapování CSV sloupců
    public class ProductMap : ClassMap<Product>
    {
        public ProductMap()
        {
            Map(m => m.OrderCode).Name("Objednací kód");
            Map(m => m.TypeDesignation).Name("Typové značení");
            Map(m => m.ProductName).Name("Název produktu");
            Map(m => m.ProductDescription).Name("Popis produktu");
            //Map(m => m.ProductDescription).Name("Popis produktu");
            Map(m => m.DiscountGroup).Name("Rabatová skupina");
            Map(m => m.BasePrice).Name("Základní cena");
            Map(m => m.Unit).Name("Měrná jednotka položky");
            Map(m => m.EAN).Name("EAN");
            Map(m => m.VolumeGross).Name("Objem (brutto) [dm3]");
            Map(m => m.WeightGross).Name("Hmotnost (brutto) [kg]");
            Map(m => m.HeightGross).Name("Výška (brutto) [mm]");
            Map(m => m.WidthGross).Name("Šířka (brutto) [mm]");
            Map(m => m.LengthGross).Name("Délka (brutto) [mm]");
            Map(m => m.CustomsCode).Name("Celní kód (intrastat)");
            Map(m => m.CountryOfOrigin).Name("Země původu");
            Map(m => m.AL).Name("AL");
            Map(m => m.ECCN).Name("ECCN");
            Map(m => m.PiecesInBasicPackage).Name("Počet kusů v základním balení");
            Map(m => m.NonDivisibleOrderQuantity).Name("Nedělitelné Objednací množství");
        }
    }
    // Mapování CSV sloupců
    public class Product
    {
        // Základní vlastnosti z CSV
        public string OrderCode { get; set; }
        public string TypeDesignation { get; set; }
        public string ProductName { get; set; }
        public string ProductDescription { get; set; } // Původní popis pro referenci
        public string DiscountGroup { get; set; }
        public decimal BasePrice { get; set; }
        public string Unit { get; set; }
        public string EAN { get; set; }
        public double VolumeGross { get; set; }
        public double WeightGross { get; set; }
        public double HeightGross { get; set; }
        public double WidthGross { get; set; }
        public double LengthGross { get; set; }
        public string CustomsCode { get; set; }
        public string CountryOfOrigin { get; set; }
        public string AL { get; set; }
        public string ECCN { get; set; }
        public double? PiecesInBasicPackage { get; set; }
        public double NonDivisibleOrderQuantity { get; set; }

        // Rozparsované atributy z ProductDescription
        public double? CurrentInA { get; set; } // In (proud, např. 10 A)
        public string VoltageUe { get; set; } // Ue (napětí, např. AC 230 V)
        public string Characteristic { get; set; } // Charakteristika (např. B, C)
        public double? ResidualCurrentIdn { get; set; } // Idn (zbytkový proud, např. 30 mA)
        public string Poles { get; set; } // Počet pólů (např. 1+N-pól, 1pól)
        public double? WidthInModules { get; set; } // Šířka v modulech
        public double? BreakingCapacityIcn { get; set; } // Jmenovitá zkratová schopnost (např. 6 kA)
        public string Type { get; set; } // Typ (např. A, AC, A-G)
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
            .Where(p => p.ProductName.Contains("Jistič") || p.ProductName.Contains("Jističochránič"))
            .Where(p => currentInA == null || p.CurrentInA == currentInA)
            .Where(p => voltageUe == null || p.VoltageUe == voltageUe)
            .Where(p => characteristic == null || p.Characteristic == characteristic)
            .Where(p => type == null || p.Type == type)
            .ToList();
    }

    public static void Main()
    {
        string file = "OEZExportZbozi2025-05-29.csv";
        string Cesta = Path.Combine(Cesty.Data, "Jištení" , file);
        List<Product> products = LoadProductsFromCsv(Cesta);

        var Pole = products.GroupBy(x => x.TypeDesignation);
        foreach(var item in Pole) {
            Console.WriteLine(item.Key);
        }

        // Najít jističe s proudem 16 A, napětím AC 230 V a charakteristikou B
        var breakers = FindCircuitBreakers(products, currentInA: 16, voltageUe: "AC 230", characteristic: "B");
        Console.WriteLine($"Nalezeno {breakers.Count} jističů (16 A, AC 230 V, charakteristika B):");
        foreach (var breaker in breakers)
        {
            Console.WriteLine($"Kód: {breaker.OrderCode}, Název: {breaker.ProductName}, Cena: {breaker.BasePrice}");
        }

        // Najít jističe typu A-G
        var agBreakers = FindCircuitBreakers(products, type: "A-G");
        Console.WriteLine($"\nNalezeno {agBreakers.Count} jističů typu A-G:");
        foreach (var breaker in agBreakers)
        {
            Console.WriteLine($"Kód: {breaker.OrderCode}, Popis: {breaker.ProductDescription}");
        }
    }
}