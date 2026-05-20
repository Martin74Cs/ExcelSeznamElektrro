using System;
using System.Collections.Generic;
using System.Text;

namespace Aplikace.Tridy {
    public class Data {
        private static readonly Data _instance = new();

        public static Data Instance => _instance;

        public string Cesta { get; set; }

        private List<(string, string)> Seznam { get; set; } = [];

        public void Set(string klic, string hodnota) {
            int index = Seznam.FindIndex(p => p.Item1 == klic);
            if(index >= 0) {
                Seznam[index] = (klic, hodnota);
            }
            else {
                Seznam.Add((klic, hodnota));
            }
        }   

        public string Get(string klic) {
            var item = Seznam.Find(p => p.Item1 == klic);
            return item != default ? item.Item2 : null;
        }


        //skrýtí konstruktoru, aby nebylo možné vytvořit další instance třídy
        private Data() { }
    }
}
