using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Item
    {
        public int _Item__id { get; set; }
        public _Item__unit _Item__cunit { get; set; }
        public _Item__unit _Item__eunit { get; set; }
        //public string revNo { get; set; } = string.Empty;
        //public string tag { get; set; } = string.Empty;
        public string _Item__revNo { get; set; } = string.Empty;
        public string _Item__tag { get; set; } = string.Empty;
        public string _Item__name { get; set; } = string.Empty;
        public string _Item__pcs { get; set; } = string.Empty;
        public int _Item__type { get; set; }

        public List<_Item__fluid> _Item__fluid { get; set; } = [];
        public float _Item__dimensionX { get; set; }
        public float _Item__dimensionY { get; set; }
        public float _Item__dimensionZ { get; set; }
        public string _Item__material { get; set; } = string.Empty;
        public string _Item__heating { get; set; } = string.Empty;
        public string _Item__mass { get; set; } = string.Empty;
        public string _Item__insul { get; set; } = string.Empty;
        public string _Item__anchor { get; set; } = string.Empty;
        public string _Item__power { get; set; } = string.Empty;
        public string _Item__noise { get; set; } = string.Empty;
        public string _Item__note { get; set; } = string.Empty;
        public List<Item> _Item__subitem { get; set; } = [];
    }

    public class _Item__unit
    {
        public int _Unit__id { get; set; }
        public string _Unit__pfx { get; set; } = string.Empty;
        public string _Unit__num { get; set; } = string.Empty;
        public string _Unit__sfx { get; set; } = string.Empty;
        public string _Unit__name { get; set; } = string.Empty;
        public string _Unit__notes { get; set; } = string.Empty;
    }

    public class _Item__fluid
    {
        public string _Fluid__fluid { get; set; }
        public float _Fluid__volume { get; set; }
        public float _Fluid__flowrate { get; set; }
        public _Fluid__parameter _Fluid__parameter { get; set; } = new();
    }

    public class _Fluid__parameter
    {
        public string _Param__value { get; set; } = string.Empty;
        public string _Param__unit { get; set; } = string.Empty;

    }
}
