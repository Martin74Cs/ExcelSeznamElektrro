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
        public string Apid { get; set; } = string.Empty;
    }
}
