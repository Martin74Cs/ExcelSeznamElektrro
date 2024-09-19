using Aplikace.Sdilene;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Tridy
{
    public class Pump
    {
        public string[] _Pump__locatedOnPIDNumber { get; set; }
        public string _Pump__tag { get; set; } = string.Empty;
        public string[] _Pump__tagSuffix { get; set; }
        public string _Pump__tagSuffixes { get; set; }
        public string _Pump__quantity { get; set; } = string.Empty;
        public string _Pump__name { get; set; } = string.Empty;
        public string _Pump__docNo { get; set; } = string.Empty;
        public string _Pump__liquidPH { get; set; } = string.Empty;
        public string _Pump__pumpedLiquid { get; set; } = string.Empty;
        public string _Pump__mixtureProportion { get; set; } = string.Empty;
        public string _Pump__mixtureProportionNote { get; set; } = string.Empty;
        public string _Pump__operatingTemperature { get; set; } = string.Empty;
        public string _Pump__maximumOperatingTemperature { get; set; } = string.Empty;
        //public float _Pump__maximumOperatingTemperature { get; set; } 
        public string _Pump__minimumOperatingTemperature { get; set; } = string.Empty;
        //public float _Pump__densityAtOperatingTemperature { get; set; } = string.Empty;
        public float _Pump__densityAtOperatingTemperature { get; set; }
        public string _Pump__viscosityAtOperatingTemperature { get; set; } = string.Empty;
        //public float _Pump__relativeVaporDensity { get; set; } = string.Empty;
        public float _Pump__relativeVaporDensity { get; set; }
        public string _Pump__designVolumeFlowrate { get; set; } = string.Empty;
        public string _Pump__normalVolumeFlowrate { get; set; } = string.Empty;
        public string _Pump__minimumVolumeFlowrate { get; set; } = string.Empty;
        public string _Pump__maximumVolumeFlowrate { get; set; } = string.Empty;
        public string _Pump__diffHead { get; set; } = string.Empty;
        public string _Pump__diffPressure { get; set; } = string.Empty;
        public string _Pump__requiredDiffHead { get; set; } = string.Empty;
        public string _Pump__electricPowerSystem { get; set; } = string.Empty;
        public string _Pump__typeOfDrive_freqencyConverter { get; set; } = string.Empty;
        public string _Pump__estimatedAbsorbedPower { get; set; } = string.Empty;
        public string _Pump__typeOfPump { get; set; } = string.Empty;
        public string _Pump__typeOfPumpNote { get; set; } = string.Empty;
        public string _Pump__maxiumNoiseLevel { get; set; } = string.Empty;
        public string _Pump__service { get; set; } = string.Empty;
        public string _Pump__materialOfConstruction { get; set; } = string.Empty;
        public string _Pump__pumpConstructionStandards { get; set; } = string.Empty;
        public string _Pump__sealing_flushingFluid_coupling { get; set; } = string.Empty;
        public string _Pump__insulation { get; set; } = string.Empty;
        public string _Pump__mass { get; set; } = string.Empty;
        public string _Pump__location { get; set; } = string.Empty;
        public string _Pump__indoor_outdoor_underRoof { get; set; } = string.Empty;
        public string _Pump__indoor_floor { get; set; } = string.Empty;
        public string _Pump__EExProofProtection { get; set; } = string.Empty;
        public string _Pump__Seismicity { get; set; } = string.Empty;
        public string _Pump__diagramPath { get; set; } = string.Empty;
        public string _Pump__status { get; set; } = string.Empty;
        public string _Pump__isPU { get; set; } = string.Empty;

        public static List<Pump> Load(string cestaPump)
        {
            //System.Data.DataTable pokus = SouboryJson.LoadJson(cesta);
            //var pokus = SouboryJson.LoadJson<Pump>(cesta);
            var pokus = Soubory.LoadJsonEn<Pump>(cestaPump);
            Console.Write($"Celkem={pokus.Count()}");
            foreach (var item in pokus)
            {
                Console.WriteLine($"Tag={item._Pump__tag}, Patro={item._Pump__indoor_floor}");
            }
            return pokus;
        }
    }
}
