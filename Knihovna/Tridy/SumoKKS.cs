using System;
using System.Collections.Generic;

namespace Parametr.MAcad
{
    #region Logovací třídy pro plný JSON výstup

    /// <summary>
    /// Reprezentuje plný log pro jeden výkresový rám.
    /// </summary>
    public class SumoDivisionLog
    {
        public string SheetHandle { get; set; } = string.Empty;
        public string CisloListu { get; set; } = string.Empty;
        public List<ZoneLog> Zones { get; set; } = new List<ZoneLog>();
        public List<SubKksLog> UnassignedKks { get; set; } = new List<SubKksLog>();

        /// <summary>
        /// Převede plná data tohoto listu na zploštělý seznam řádků pro tabulkový tisk/export.
        /// </summary>
        public List<SumoDivisionFlatRow> ToFlatRows()
        {
            var list = new List<SumoDivisionFlatRow>();

            // 1. Zóny a v nich zařazené KKS
            foreach (var zone in Zones)
            {
                // Hlavní KKS (pokud existuje)
                if (!string.IsNullOrEmpty(zone.MainKksHandle))
                {
                    list.Add(new SumoDivisionFlatRow
                    {
                        SheetHandle = this.SheetHandle,
                        CisloListu = this.CisloListu,
                        ZoneIndex = zone.ZoneIndex,
                        Rozvadec = zone.RozvadecText,
                        HlavniKKS = zone.MainKksValue,
                        KksHandle = zone.MainKksHandle,
                        KksHodnota = zone.MainKksValue
                    });
                }

                // Podružné KKS
                foreach (var sub in zone.SubordinateKks)
                {
                    list.Add(new SumoDivisionFlatRow
                    {
                        SheetHandle = this.SheetHandle,
                        CisloListu = this.CisloListu,
                        ZoneIndex = zone.ZoneIndex,
                        Rozvadec = zone.RozvadecText,
                        HlavniKKS = zone.MainKksValue,
                        KksHandle = sub.Handle,
                        KksHodnota = sub.NewValue
                    });
                }
            }

            // 2. Nezařazené KKS (pokud existují)
            foreach (var un in UnassignedKks)
            {
                list.Add(new SumoDivisionFlatRow
                {
                    SheetHandle = this.SheetHandle,
                    CisloListu = this.CisloListu,
                    ZoneIndex = -1, // -1 značí nezařazené
                    Rozvadec = "(Mimo zóny)",
                    HlavniKKS = "-",
                    KksHandle = un.Handle,
                    KksHodnota = un.OriginalValue
                });
            }

            return list;
        }
    }

    /// <summary>
    /// Reprezentuje plný log pro jednu uzavřenou zónu v rámu.
    /// </summary>
    public class ZoneLog
    {
        public int ZoneIndex { get; set; }
        public string MainKksValue { get; set; } = string.Empty;
        public string MainKksHandle { get; set; } = string.Empty;
        public string MainKksPosition { get; set; } = string.Empty;
        public List<KksAttributeLog> MainKksAttributes { get; set; } = new List<KksAttributeLog>();
        public string RozvadecText { get; set; } = string.Empty;
        public string RozvadecTextPosition { get; set; } = string.Empty;
        public List<SubKksLog> SubordinateKks { get; set; } = new List<SubKksLog>();
    }

    /// <summary>
    /// Reprezentuje plný log pro jeden podružný KKS blok.
    /// </summary>
    public class SubKksLog
    {
        public string Handle { get; set; } = string.Empty;
        public string Position { get; set; } = string.Empty;
        public string OriginalValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
        public bool Updated { get; set; }
        public List<KksAttributeLog> Attributes { get; set; } = new List<KksAttributeLog>();
    }

    /// <summary>
    /// Reprezentuje log pro atribut KKS bloku (Tag a Hodnota).
    /// </summary>
    public class KksAttributeLog
    {
        public string Tag { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    #endregion

    #region Zjednodušené logovací třídy pro export

    /// <summary>
    /// Reprezentuje zjednodušený log pro jeden výkresový rám.
    /// </summary>
    public class SumoDivisionSimpleLog
    {
        public string SheetHandle { get; set; } = string.Empty;
        public string CisloListu { get; set; } = string.Empty;
        public List<ZoneSimpleLog> Zones { get; set; } = new List<ZoneSimpleLog>();
    }

    /// <summary>
    /// Reprezentuje zjednodušený log pro jednu zónu.
    /// </summary>
    public class ZoneSimpleLog
    {
        public int ZoneIndex { get; set; }
        public string Rozvadec { get; set; } = string.Empty;
        public string HlavniKKS { get; set; } = string.Empty;
        public List<KksSimpleLog> KksList { get; set; } = new List<KksSimpleLog>();
    }

    /// <summary>
    /// Reprezentuje zjednodušené informace o KKS bloku.
    /// </summary>
    public class KksSimpleLog
    {
        public string Handle { get; set; } = string.Empty;
        public string Hodnota { get; set; } = string.Empty;
    }

    #endregion

    #region Zploštělé logovací třídy pro tisk a tabulky

    /// <summary>
    /// Reprezentuje jeden plochý řádek (záznam) zploštělé struktury pro tisk nebo export do tabulky.
    /// </summary>
    public class SumoDivisionFlatRow
    {
        public string SheetHandle { get; set; } = string.Empty;
        public string CisloListu { get; set; } = string.Empty;
        public int ZoneIndex { get; set; }
        public string Rozvadec { get; set; } = string.Empty;
        public string HlavniKKS { get; set; } = string.Empty;
        public string KksHandle { get; set; } = string.Empty;
        public string KksHodnota { get; set; } = string.Empty;
    }

    /// <summary>
    /// Statická třída pro extension metody nad logy.
    /// </summary>
    public static class SumoDivisionExtensions
    {
        /// <summary>
        /// Převede seznam plných logů na zploštělý seznam řádků pro tabulkový tisk/export.
        /// </summary>
        public static List<SumoDivisionFlatRow> ToFlatRows(this IEnumerable<SumoDivisionLog> logs)
        {
            var result = new List<SumoDivisionFlatRow>();
            if (logs == null) return result;

            foreach (var log in logs)
            {
                result.AddRange(log.ToFlatRows());
            }
            return result;
        }
    }

    #endregion
}
