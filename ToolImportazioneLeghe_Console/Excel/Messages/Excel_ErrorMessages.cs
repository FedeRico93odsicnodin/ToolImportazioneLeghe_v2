using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Messages
{
    /// <summary>
    /// Tutti i messaggi di errore che si hanno durante la lettura / validazione foglio / recupero informazioni / validazione informazioni 
    /// per il foglio excel corrente nelle diverse versioni per i formati e fogli disponibili
    /// </summary>
    public static class Excel_ErrorMessages
    {
        /// <summary>
        /// Messaggi di errore per la prima tipologia di foglio per il primo formato
        /// </summary>
        public static class Formato1_Foglio1_Leghe
        {

            /// <summary>
            /// Errore nella lettura di una proprieta obbligatoria per la lega correntemente in lettura, manca il valore e la lega non verrà presa in considerazione nelle analisi successive
            /// da formattare con 
            /// 1) numero di riga in cui l'informazione corrente è mancante
            /// 2) la proprieta non valorizzata per la lega corrente
            /// </summary>
            public static string ERRORE_MANDATORYPROPERTYMANCANTE_LEGA = "ERRORE - riga {0}: nessuna informazione valorizzata per la proprietà obbligatoria '{1}', la seguente lega presa in considerazione nelle analisi successive.\n";

            /// <summary>
            /// Errore fatale: questo errore non permette di continuare con l'iterazione per il foglio excel corrente 
            /// </summary>
            public static string ERRORE_NESSUNA_INFORMAZIONE_LETTA = "ERRORE: non è stata letta alcuna informazione per il foglio Excel di LEGHE '{0}'.\n";
        }


        /// <summary>
        /// Messaggi di errore per la seconda tipologia di foglio per il primo formato
        /// </summary>
        public static class Formato1_Foglio2_Concentrazioni
        {
            /// <summary>
            /// Errore di mancata lettura del nome per il quadrante delle concentrazioni correntemente in analisi per la situazione corrente
            /// </summary>
            public static string ERRORE_NOMEMATERIALELETTURAQUADRANTEVUOTO = "ERRORE - riga {0}, colonna {1}: nessuna informazione attribuita al nome per il MATERIALE per questo quadrante.\n";


            /// <summary>
            /// Errore durante l'analisi delle proprieta obbligatorie relative al quadrante di concentrazione corrente e per quanto riguarda le 
            /// proprieta relative alle diverse concentrazioni applicate per la lega corrente 
            /// </summary>
            public static string ERRORE_MANDATORYPROPERTYMANCANTE_CONCENTRAZIONI = "ERRORE - riga {0}: nessuna inforamzione valorizzata per la proprieta obbligatoria '{1}', le concentrazioni per la lega non verranno prese in considerazione per le analisi successive.\n";
        }


        /// <summary>
        /// Messaggi di errore per la prima tipologia di foglio e il secondo formato disponibile
        /// </summary>
        public static class Formato2_Foglio1_LegheConcentrazioni
        {

        }
    }
}
