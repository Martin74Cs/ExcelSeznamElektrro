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
        public string[] Pump__locatedOnPIDNumber { get; set; } = [];
        public string Pump__tag { get; set; } = string.Empty;
        public string[] Ump__tagSuffix { get; set; } = [];
        public string Ump__tagSuffixes { get; set; } = string.Empty;
        public string Pump__quantity { get; set; } = string.Empty;
        public string Pump__name { get; set; } = string.Empty;
        public string Pump__docNo { get; set; } = string.Empty;
        public string Pump__liquidPH { get; set; } = string.Empty;
        public string Pump__pumpedLiquid { get; set; } = string.Empty;
        public string Pump__mixtureProportion { get; set; } = string.Empty;
        public string Pump__mixtureProportionNote { get; set; } = string.Empty;
        public string Pump__operatingTemperature { get; set; } = string.Empty;
        public string Pump__maximumOperatingTemperature { get; set; } = string.Empty;
        //public float _Pump__maximumOperatingTemperature { get; set; } 
        public string Pump__minimumOperatingTemperature { get; set; } = string.Empty;
        //public float _Pump__densityAtOperatingTemperature { get; set; } = string.Empty;
        public float Pump__densityAtOperatingTemperature { get; set; }
        public string Pump__viscosityAtOperatingTemperature { get; set; } = string.Empty;
        //public float _Pump__relativeVaporDensity { get; set; } = string.Empty;
        public float Pump__relativeVaporDensity { get; set; }
        public string Pump__designVolumeFlowrate { get; set; } = string.Empty;
        public string Pump__normalVolumeFlowrate { get; set; } = string.Empty;
        public string Pump__minimumVolumeFlowrate { get; set; } = string.Empty;
        public string Pump__maximumVolumeFlowrate { get; set; } = string.Empty;
        public string Pump__diffHead { get; set; } = string.Empty;
        public string Pump__diffPressure { get; set; } = string.Empty;
        public string Pump__requiredDiffHead { get; set; } = string.Empty;
        public string Pump__electricPowerSystem { get; set; } = string.Empty;
        public string Pump__typeOfDrive_freqencyConverter { get; set; } = string.Empty;
        public string Pump__estimatedAbsorbedPower { get; set; } = string.Empty;
        public string Pump__typeOfPump { get; set; } = string.Empty;
        public string Pump__typeOfPumpNote { get; set; } = string.Empty;
        public string Pump__maxiumNoiseLevel { get; set; } = string.Empty;
        public string Pump__service { get; set; } = string.Empty;
        public string Pump__materialOfConstruction { get; set; } = string.Empty;
        public string Pump__pumpConstructionStandards { get; set; } = string.Empty;
        public string Pump__sealing_flushingFluid_coupling { get; set; } = string.Empty;
        public string Pump__insulation { get; set; } = string.Empty;
        public string Pump__mass { get; set; } = string.Empty;
        public string Pump__location { get; set; } = string.Empty;
        public string Pump__indoor_outdoor_underRoof { get; set; } = string.Empty;
        public string Pump__indoor_floor { get; set; } = string.Empty;
        public string Pump__EExProofProtection { get; set; } = string.Empty;
        public string Pump__Seismicity { get; set; } = string.Empty;
        public string Pump__diagramPath { get; set; } = string.Empty;
        public string Pump__status { get; set; } = string.Empty;
        public string Pump__isPU { get; set; } = string.Empty;

        public static List<Pump> Load(string cestaPump)
        {
            //System.Data.DataTable pokus = SouboryJson.LoadJson(cesta);
            //var pokus = SouboryJson.LoadJson<Pump>(cesta);
            var pokus = Soubory.LoadJsonEn<Pump>(cestaPump);
            Console.Write($"Celkem={pokus.Count}");
            foreach (var item in pokus)
            {
                Console.WriteLine($"Tag={item.Pump__tag}, Patro={item.Pump__indoor_floor}");
            }
            return pokus;
        }
    }
}
