using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Knihovna.Shared.Tridy
{
    /// <summary>
    /// Poskytuje přístup k databázi kabelů (Cu a Al) načtené z JSON souborů.
    /// </summary>
    public static class KabelDatabaze
    {
        private static readonly object Zamek = new();
        private static List<Kabel> cuKabely = [];
        private static List<Kabel> alKabely = [];
        private static bool inicializovano = false;

        /// <summary>Seznam načtených měděných kabelů.</summary>
        public static List<Kabel> CuKabely
        {
            get
            {
                lock (Zamek)
                {
                    return cuKabely;
                }
            }
        }

        /// <summary>Seznam načtených hliníkových kabelů.</summary>
        public static List<Kabel> AlKabely
        {
            get
            {
                lock (Zamek)
                {
                    return alKabely;
                }
            }
        }

        /// <summary>
        /// Inicializuje databázi kabelů ze zadaných cest k JSON souborům.
        /// </summary>
        public static void Inicializuj(string cestaCu, string cestaAl)
        {
            lock (Zamek)
            {
                if (inicializovano) return;

                try
                {
                    if (File.Exists(cestaCu))
                    {
                        string json = File.ReadAllText(cestaCu);
                        cuKabely = JsonConvert.DeserializeObject<List<Kabel>>(json) ?? [];
                    }

                    if (File.Exists(cestaAl))
                    {
                        string json = File.ReadAllText(cestaAl);
                        alKabely = JsonConvert.DeserializeObject<List<Kabel>>(json) ?? [];
                    }

                    inicializovano = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chyba při načítání databáze kabelů: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Vyhledá kabel v databázi na základě typu, počtu žil a průřezu.
        /// </summary>
        public static Kabel? NajdiKabel(string typ, string pocetZil, string prurez)
        {
            if (string.IsNullOrWhiteSpace(pocetZil) || string.IsNullOrWhiteSpace(prurez))
            {
                return null;
            }

            // Určení materiálu - začíná na 'A'/'a' = Al, jinak Cu
            bool jeHlinik = typ.Trim().StartsWith("A", StringComparison.OrdinalIgnoreCase);
            List<Kabel> databaze = jeHlinik ? AlKabely : CuKabely;

            if (databaze.Count == 0)
            {
                return null;
            }

            // Normalizace počtu žil (např. "3G" -> "3", "3G2.5" -> "3")
            string cistyPocetZil = pocetZil.Trim();
            int indexG = cistyPocetZil.IndexOf('G', StringComparison.OrdinalIgnoreCase);
            if (indexG > 0)
            {
                cistyPocetZil = cistyPocetZil.Substring(0, indexG).Trim();
            }
            int indexX = cistyPocetZil.IndexOf('x', StringComparison.OrdinalIgnoreCase);
            if (indexX > 0)
            {
                cistyPocetZil = cistyPocetZil.Substring(0, indexX).Trim();
            }

            // Normalizace průřezu (např. "1,5" -> 1.5)
            string cistyPrurez = prurez.Trim().Replace(',', '.');
            if (!double.TryParse(cistyPrurez, System.Globalization.NumberStyles.Any, 
                System.Globalization.CultureInfo.InvariantCulture, out double prurezHodnota))
            {
                return null;
            }

            // Normalizace typu kabelu (např. "CYKY-J" -> "CYKY")
            string cistyTyp = typ.Trim();
            int indexPomlcky = cistyTyp.IndexOf('-');
            if (indexPomlcky > 0)
            {
                cistyTyp = cistyTyp.Substring(0, indexPomlcky).Trim();
            }

            // Krok 1: Hledáme kabely se shodným průřezem a počtem žil
            var shodaRozmer = databaze.Where(k => 
                k.SLmm2 == prurezHodnota && 
                (k.Deleni == cistyPocetZil || k.Deleni.Trim() == cistyPocetZil)).ToList();

            if (shodaRozmer.Count == 0)
            {
                return null;
            }

            // Krok 2: Hledáme podle názvu
            var shodaNazev = shodaRozmer.Where(k => 
                k.Name.Contains(cistyTyp, StringComparison.OrdinalIgnoreCase) || 
                k.Označení.Contains(cistyTyp, StringComparison.OrdinalIgnoreCase)).ToList();

            if (shodaNazev.Count > 0)
            {
                // Pokud máme shodu, vezmeme první (nebo nejbližší)
                return shodaNazev.First();
            }

            // Fallback: Pokud nenašli podle názvu, vrátíme první se shodným rozměrem
            return shodaRozmer.First();
        }
    }
}
