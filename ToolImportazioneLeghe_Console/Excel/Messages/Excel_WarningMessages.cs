using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Messages
{
    /// <summary>
    /// Tutti i messaggi di warnings che si hanno durante la lettura / validazione foglio / recupero informazioni / validazione informazioni 
    /// per il foglio excel corrente nelle diverse versioni per i formati e fogli disponibili
    /// </summary>
    public class Excel_WarningMessages
    {
        /// <summary>
        /// Messaggi di warings per la prima tipologia di foglio per il primo formato
        /// </summary>
        public static class Formato1_Foglio1_Leghe
        {
            /// <summary>
            /// Segnalazione di mancanza di un valore per la proprieta opzionale di lega 
            /// 1) riga nella quale la proprieta è mancante
            /// 2) nome della proprieta opzionale per la quale manca il valore
            /// </summary>
            public static string WARNING_MANCANZAVALOREPERPROPRIETAOPZIONALE_LEGA = "WARNING - riga {0}: mancanza del valore per la poprietà opzionale '{1}'";
        }


        /// <summary>
        /// Messaggi di warings per la seconda tipologia di foglio per il primo formato
        /// </summary>
        public static class Formato1_Foglio2_Concentrazioni
        {
            /// <summary>
            /// Segnalazione della mancanza di valore per una proprieta opzionale di concentrazione di lettura per una particolare riga all'interno del foglio di tipo 2 
            /// per il formato 1 e di cui viene passato anche il nome corrispondente 
            /// </summary>
            public static string WARNING_MANCANZAVALOREPERPROPRIETAOPZIONALE_CONCENTRAZIONI = "WARNING - riga {0}: mancanza del valore per la proprieta opzionale '{1}' di concentrazione";
        }


        /// <summary>
        /// Messaggi di warnings per la prima tipologia di foglio e il secondo formato disponibile
        /// </summary>
        public static class Formato2_Foglio1_LegheConcentrazioni
        {
            /// <summary>
            /// Segnalazione di mancata lettura per una proprieta opzionale rispetto al foglio corrente di secondo formato e iterazione su una particolare lega
            /// </summary>
            public static string WARNING_MANCATALETTURAPROPRIETAOPZIONALE_LEGA = "WARNING - riga {0}, colonna {1}: mancata lettura della proprieta '{2}' OPZIONALE per le informazioni di lega";


            /// <summary>
            /// Segnalazione di mancata lettura per una proprieta opzionale rispetto al foglio corrente di secondo formato e iterazione su proprieta d un determinato elemento
            /// </summary>
            public static string WARNING_MANCATALETTURAPROPRIETAOPZIONALE_CONCENTRAZIONI = "WARNING - riga {0}, colonna {1}: mancata lettura della proprieta '{2}' OPZIONALE per le informazioni di concentrazione";
        }
    }
}
