using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Item
    {
        public int Id { get; set; }
        public Unit Cunit { get; set; } = new();
        public Unit Munit { get; set; } = new();    
        
        //public string revNo { get; set; } = string.Empty;
        //public string tag { get; set; } = string.Empty;
        
        public string Tag { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Pcs { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;

        public List<Fluids> Fluid { get; set; } = [];
        public float DimensionX { get; set; }
        public float DimensionY { get; set; }
        public float DimensionZ { get; set; }
        public string Material { get; set; } = string.Empty;
        public string Heating { get; set; } = string.Empty;
        public string Mass { get; set; } = string.Empty;
        public string Insul { get; set; } = string.Empty;
        public string Anchor { get; set; } = string.Empty;
        public string Power { get; set; } = string.Empty;
        public string Noise { get; set; } = string.Empty;
        public string Note { get; set; } = string.Empty;
        public List<Item> Subitem { get; set; } = [];
    }

    public class Unit
    {
        public int Id { get; set; }
        public string Pfx { get; set; } = string.Empty;
        public string Num { get; set; } = string.Empty;
        public string Sfx { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Notes { get; set; } = string.Empty;
    }

    public class Fluids
    {
        public string Fluid { get; set; }
        public float Volume { get; set; }
        public float Flowrate { get; set; }
        public Parameter Parameter { get; set; } = new();
    }

    public class Parameter
    {
        public string Value { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;

    }
}
