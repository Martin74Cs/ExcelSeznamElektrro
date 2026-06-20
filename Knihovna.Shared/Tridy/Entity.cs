namespace Knihovna.Tridy
{
    /// <summary>
    /// Základní třída reprezentující entitu s identifikátorem.
    /// </summary>
    public class Entity
    {
        /// <summary>
        /// Unikátní identifikátor entity.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Externí/API identifikátor entity.
        /// </summary>
        public string Apid { get; set; }

        /// <summary>
        /// Inicializuje novou instanci třídy <see cref="Entity"/> a vygeneruje náhodný 8místný Apid.
        /// </summary>
        public Entity()
        {
            Apid = GenerujApid();
        }

        /// <summary>
        /// Generuje náhodný 8místný alfanumerický řetězec.
        /// </summary>
        private static string GenerujApid()
        {
            const string znaky = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            char[] vysledek = new char[8];
            for (int i = 0; i < 8; i++)
            {
                vysledek[i] = znaky[Random.Shared.Next(znaky.Length)];
            }
            return new string(vysledek);
        }
    }
}
