using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplikace.Sdilene
{
    public class Klavesnice
    {
        public static ConsoleKeyInfo? VstupTimeoutem(string Text, int timeoutMs=1000)
        {
            Console.Write(Text);
            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                if (Console.KeyAvailable)
                    return Console.ReadKey(intercept: true); // intercept = true => nezobrazí znak
                Thread.Sleep(50); // trochu šetrnější k CPU
            }
            return null; 
        }
        public static bool VstupAnoTimeoutem(string Text, int timeoutMs=1000)
        {
            Console.Write(Text);
            var key = VstupTimeoutem(Text, timeoutMs);
            if (key.Value.Key == ConsoleKey.A)
                return true;
            return false;
        }
    }
}
