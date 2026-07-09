using Knihovna.Shared.Tridy;
using System.ComponentModel;

namespace Knihovna.Tridy
{
    /// <summary>
    /// Reprezentuje zobrazení jednoho kabelu (trasy) provázaného se zařízením pro zobrazení v Gridu.
    /// </summary>
    public class KabelRowView
    {
        /// <summary>
        /// Odkaz na mateřské zařízení.
        /// </summary>
        [Browsable(false)]
        public Zarizeni Zarizeni { get; }

        /// <summary>
        /// Odkaz na podkladovou trasu (kabel).
        /// </summary>
        [Browsable(false)]
        public Trasa Trasa { get; }

        /// <summary>
        /// Inicializuje novou instanci třídy <see cref="KabelRowView"/>.
        /// </summary>
        /// <param name="zarizeni">Mateřské zařízení.</param>
        /// <param name="trasa">Podkladová trasa kabelu.</param>
        public KabelRowView(Zarizeni zarizeni, Trasa trasa)
        {
            Zarizeni = zarizeni;
            Trasa = trasa;
        }

        /// <summary>
        /// Označení (Tag) zařízení.
        /// </summary>
        [DisplayName("Zařízení (Tag)")]
        [ReadOnly(true)]
        public string ZarizeniTag => Zarizeni.Tag;

        /// <summary>
        /// Předmět zařízení.
        /// </summary>
        [DisplayName("Předmět")]
        [ReadOnly(true)]
        public string ZarizeniPredmet => Zarizeni.Predmet;

        /// <summary>
        /// Jméno nebo popis zařízení.
        /// </summary>
        [DisplayName("Popis zařízení")]
        [ReadOnly(true)]
        public string ZarizeniPopis => Zarizeni.Popis;

        /// <summary>
        /// Označení kabelu (např. WL 01).
        /// </summary>
        [DisplayName("Označení")]
        public string Oznaceni
        {
            get => Trasa.Oznaceni;
            set
            {
                Trasa.Oznaceni = value;
                // Aktualizace přehledu kabelů na zařízení
                Zarizeni.NotifyPropertyChanged(nameof(Zarizeni.KabelyPrehled));
            }
        }

        /// <summary>
        /// Typ kabelu (např. CYKY-J).
        /// </summary>
        [DisplayName("Kabel (Typ)")]
        public string Kabel
        {
            get => Trasa.Kabel;
            set => Trasa.Kabel = value;
        }

        /// <summary>
        /// Počet žil kabelu.
        /// </summary>
        [DisplayName("Počet žil")]
        public string PocetZil
        {
            get => Trasa.PocetZil;
            set => Trasa.PocetZil = value;
        }

        /// <summary>
        /// Průřez vodičů v mm2.
        /// </summary>
        [DisplayName("Průřez [mm2]")]
        public string Prurezmm2
        {
            get => Trasa.Prurezmm2;
            set => Trasa.Prurezmm2 = value;
        }

        /// <summary>
        /// Délka kabelu v metrech.
        /// </summary>
        [DisplayName("Délka [m]")]
        public string Delka
        {
            get => Trasa.Delka;
            set => Trasa.Delka = value;
        }

        /// <summary>
        /// Popis nebo poznámka ke kabelu.
        /// </summary>
        [DisplayName("Svorka")]
        public string Svorka
        {
            get => Trasa.Svorka;
            set => Trasa.Svorka = value;
        }

        /// <summary>
        /// Popis nebo poznámka ke kabelu.
        /// </summary>
        [DisplayName("Popis / Poznámka")]
        public string Popis
        {
            get => Trasa.Popis;
            set => Trasa.Popis = value;
        }

        /// <summary>
        /// Označení rozvaděče, ze kterého je kabel napájen.
        /// </summary>
        [DisplayName("Rozvaděč")]
        [ReadOnly(true)]
        public string RozvadecAll => Trasa.RozvadecAll;

        /// <summary>
        /// Proudové zatížení kabelu (např. Iz = 22/34 A).
        /// </summary>
        [DisplayName("Zatížení (Iz)")]
        [ReadOnly(true)]
        public string ProudZatizeni => Trasa.KabelData?.Proud ?? string.Empty;
    }
}
