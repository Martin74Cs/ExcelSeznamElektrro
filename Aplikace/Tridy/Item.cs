using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Item
    {
        public int id { get; set; }
        public Unit cunit { get; set; }
        public Unit munit { get; set; }
        //public string revNo { get; set; } = string.Empty;
        //public string tag { get; set; } = string.Empty;
        //public string revNo { get; set; } = string.Empty;
        public string tag { get; set; } = string.Empty;
        public string name { get; set; } = string.Empty;
        public string pcs { get; set; } = string.Empty;
        public string type { get; set; }

        public List<Fluid> fluid { get; set; } = [];
        public float dimensionX { get; set; }
        public float dimensionY { get; set; }
        public float dimensionZ { get; set; }
        public string material { get; set; } = string.Empty;
        public string heating { get; set; } = string.Empty;
        public string mass { get; set; } = string.Empty;
        public string insul { get; set; } = string.Empty;
        public string anchor { get; set; } = string.Empty;
        public string power { get; set; } = string.Empty;
        public string noise { get; set; } = string.Empty;
        public string note { get; set; } = string.Empty;
        public List<Item> subitem { get; set; } = [];
    }

    public class Unit
    {
        public int id { get; set; }
        public string pfx { get; set; } = string.Empty;
        public string num { get; set; } = string.Empty;
        public string sfx { get; set; } = string.Empty;
        public string name { get; set; } = string.Empty;
        public string notes { get; set; } = string.Empty;
    }

    public class Fluid
    {
        public string fluid { get; set; }
        public float volume { get; set; }
        public float flowrate { get; set; }
        public Parameter parameter { get; set; } = new();
    }

    public class Parameter
    {
        public string value { get; set; } = string.Empty;
        public string unit { get; set; } = string.Empty;

    }
}
