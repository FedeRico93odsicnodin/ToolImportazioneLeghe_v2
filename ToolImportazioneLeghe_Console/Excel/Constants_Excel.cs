using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel
{
    /// <summary>
    /// Inserimento di tutte le costati di lettura del foglio excel, che sia per il primo o il secondo formato e per i diversi 
    /// fogli disponibili all'interno del file 
    /// </summary>
    public static class Constants_Excel
    {
        #region PARAMETRI PER LA LETTURA DEL FORMATO 1 PRESO DAL DATABASE DI LEGHE

        /// <summary>
        /// Indica se il foglio correntemente in analisi è relativo alla lettura delle informazioni di lega 
        /// piuttosto che dei quadranti per le diverse concentrazioni
        /// </summary>
        public enum TipologiaFoglio_Format1
        {
            NotDefined = 0,
            FoglioLeghe = 1,
            FoglioConcentrazioni = 2
        }


        /// <summary>
        /// Esito della validazione per il formato excel 1 - le informazioni lette e validate possono rientrare in una della seguente casistica 
        /// - per la prima tipologia di lettura si prosegue senza chiedere nessuna conferma a user
        /// - per la seconda tipologia si chiede conferma allo user anche in base alla modalita impostata di sovrascrittura o accodamento informazioni
        /// - per la terza tipologia si ferma il tool
        /// </summary>
        public enum ValidazioneFoglio
        {
            LetturaCompleta = 1,
            LetturaParziale = 2,
            NessunaCorrispondenza = 3
        }

        #endregion
    }
}
