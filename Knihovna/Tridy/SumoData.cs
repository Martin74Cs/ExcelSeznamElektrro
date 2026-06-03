
namespace Parametr.MAcad
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

}
