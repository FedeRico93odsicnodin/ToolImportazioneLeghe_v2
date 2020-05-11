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
            /// Errore fatale: questo errore non permette di continuare con l'iterazione per il foglio excel corrente 
            /// </summary>
            public static string ERRORE_NESSUNA_INFORMAZIONE_LETTA = "ERRORE: non è stata letta alcuna informazione per il foglio Excel di LEGHE '{0}'";
        }


        /// <summary>
        /// Messaggi di errore per la seconda tipologia di foglio per il primo formato
        /// </summary>
        public static class Formato1_Foglio2_Concentrazioni
        {

        }


        /// <summary>
        /// Messaggi di errore per la prima tipologia di foglio e il secondo formato disponibile
        /// </summary>
        public static class Formato2_Foglio1_LegheConcentrazioni
        {

        }
    }
}
