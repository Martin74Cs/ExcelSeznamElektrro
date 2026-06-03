namespace Knihovna.Tridy
{

    public class SumoResult
    {
        public string SumoHandle { get; set; } = string.Empty;
        public int RazitkaCount { get; set; }
        public string RazitkoList { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public int KksCount { get; set; }
        public List<KksItem> Kks { get; set; } = [];
    }

    public class KksItem
    {
        public List<AttrItem> AtributySPomlckou { get; set; } = [];
        public string ZmenaPred { get; set; } = string.Empty;
        public string ZmenaPo { get; set; } = string.Empty;
    }

    public class AttrItem
    {
        public string Tag { get; set; } = string.Empty;
        public string Hodnota { get; set; } = string.Empty;
    }


    public class FlatRow
    {
        public string Sumo { get; set; } = string.Empty;
        public string Razitko { get; set; } = string.Empty;
        public string Kks { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public string Hodnota { get; set; } = string.Empty;
        public string ZmenaPred { get; set; } = string.Empty;
        public string ZmenaPo { get; set; } = string.Empty;


        public static List<FlatRow> Flatten(List<SumoResult> data)
        {
            var result = new List<FlatRow>();

            foreach (var sumo in data)
            {
                foreach (var kks in sumo.Kks)
                {
                    // pokud nejsou atributy → pořád vytvoř řádek
                    if (kks.AtributySPomlckou == null || kks.AtributySPomlckou.Count == 0)
                    {
                        result.Add(new FlatRow
                        {
                            Sumo = sumo.SumoHandle,
                            Razitko = sumo.RazitkoList,
                            Kks = sumo.Text,
                            Tag = "",
                            Hodnota = "",
                            ZmenaPred = kks.ZmenaPred,
                            ZmenaPo = kks.ZmenaPo
                        });

                        continue;
                    }

                    // 🔥 hlavní rozbalení
                    foreach (var atr in kks.AtributySPomlckou)
                    {
                        result.Add(new FlatRow
                        {
                            Sumo = sumo.SumoHandle,
                            Razitko = sumo.RazitkoList,
                            Kks = sumo.Text,
                            Tag = atr.Tag,
                            Hodnota = atr.Hodnota,
                            ZmenaPred = kks.ZmenaPred,
                            ZmenaPo = kks.ZmenaPo
                        });
                    }
                }
            }

            return result;
        }


    }


}
