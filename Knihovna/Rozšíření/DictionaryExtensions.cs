using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Rozšíření
{
using System.Collections.Generic;

    public static class DictionaryExtensions
    {
        /// <summary>
        /// Přidá prvky ze zdrojového slovníku do cílového. Může přepisovat existující hodnoty.
        /// </summary>
        /// <typeparam name="TKey">Typ klíče.</typeparam>
        /// <typeparam name="TValue">Typ hodnoty.</typeparam>
        /// <param name="target">Cílový slovník, do kterého se budou přidávat prvky.</param>
        /// <param name="source">Zdrojový slovník, ze kterého se budou prvky přidávat.</param>
        /// <param name="overwriteExisting">Pokud je true, existující hodnoty se přepíší.</param>
        public static IDictionary<TKey, TValue> AddRange<TKey, TValue>(
            this IDictionary<TKey, TValue> target,
            IDictionary<TKey, TValue> source,
            bool overwriteExisting = false)
        {
            foreach (var kvp in source)
            {
                if (overwriteExisting || !target.ContainsKey(kvp.Key))
                {
                    target[kvp.Key] = kvp.Value;
                }
            }
            return target;
        }
    }
}
