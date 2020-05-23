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
            /// Segnalazione di mancato riconoscimento di una proprieta di header opzionale per il foglio relativo alle proprieta di lega per il primo formato excel
            /// </summary>
            public static string WARNING_MANCATORICONOSCIMENTOPROPRIETAHEADER_LEGA = "WARNING - non sono riuscito a trovate la seguente proprieta '{0}' OPZIONALE per il foglio '{1}' di proprieta di lega per il primo formato";


            /// <summary>
            /// Segnalazione di mancanza di informazione per una intera riga in lettura per l'istanza di recupero leghe corrente dal primo foglio per il primo formato
            /// </summary>
            public static string WARNING_HOTROVATOUNARIGACOMPLETAMENTEVUOTA_LEGA = "WARNING - la riga {0} non contiene informazioni per l'istanza di recupero per le leghe correnti";


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

            #region RICONOSCIMENTO TIPOLOGIA FOGLIO 

            /// <summary>
            /// Segnalazione di mancato riconoscimento per una proprieta opzionale di header per il quadrante delle concentrazioni e il primo formato excel disponibile
            /// </summary>
            public static string WARNING_MANCATORICONOSCIMENTOPROPRIETAOPZIONALIQUADRANTE = "WARNING - riga {0}: mancato riconoscimento per l'header della proprieta opzionale '{1}'";
            
            #endregion 

            
            #region RECUPERO INFORMAZIONI - VALIDAZIONE 1

            /// <summary>
            /// Segnalazione della mancanza di valore per una proprieta opzionale di concentrazione di lettura per una particolare riga all'interno del foglio di tipo 2 
            /// per il formato 1 e di cui viene passato anche il nome corrispondente 
            /// </summary>
            public static string WARNING_MANCANZAVALOREPERPROPRIETAOPZIONALE_CONCENTRAZIONI = "WARNING - riga {0}: mancanza del valore per la proprieta opzionale '{1}' di concentrazione";

            #endregion
        }


        /// <summary>
        /// Messaggi di warnings per la prima tipologia di foglio e il secondo formato disponibile
        /// </summary>
        public static class Formato2_Foglio1_LegheConcentrazioni
        {
            #region WARNINGS SU RICONOSCIMENTO SECONDO FORMATO

            /// <summary>
            /// Segnalazione per la mancata lettura di una proprieta opzionale di lega, comunque non invalidante per le successive analisi e valorizzazioni rispetto ai valori 
            /// contenuti nel foglio di secondo tipo corrente 
            /// </summary>
            public static string WARNING_MANCATORICONOSCIMENTOPROPRIETAOPZIONALE_LEGA = "WARNING - riga {0}: mancata lettura della seguente proprieta OPZIONALE DI LEGA '{1}'";


            /// <summary>
            /// Segnalazione di mancata lettura di un header opzionale di concentrazione, la mancata lettura di questo header non è invalidante ma è comunque segnalata nei messaggi 
            /// finali per il foglio e l'iterazione correnti
            /// </summary>
            public static string WARNING_MANCATORICONOSCIMENTOPROPRIETAOPZIONALE_CONCENTRAZIONI = "WARNING - riga {0}: mancata lettura della seguente proprieta OPZIONALE DI CONCENTRAZIONI '{1}'";

            #endregion


            #region WARNINGS SU LETTURA PROPRIETA - VALIDAZIONE 1

            /// <summary>
            /// Segnalazione di mancata lettura per una proprieta opzionale rispetto al foglio corrente di secondo formato e iterazione su una particolare lega
            /// </summary>
            public static string WARNING_MANCATALETTURAPROPRIETAOPZIONALE_LEGA = "WARNING - riga {0}, colonna {1}: mancata lettura della proprieta '{2}' OPZIONALE per le informazioni di lega";


            /// <summary>
            /// Segnalazione di mancata lettura per una proprieta opzionale rispetto al foglio corrente di secondo formato e iterazione su proprieta d un determinato elemento
            /// </summary>
            public static string WARNING_MANCATALETTURAPROPRIETAOPZIONALE_CONCENTRAZIONI = "WARNING - riga {0}, colonna {1}: mancata lettura della proprieta '{2}' OPZIONALE per le informazioni di concentrazione";


            /// <summary>
            /// Segnalazione relativa alla mancata lettura completa di tutte le proprieta per un certo elemento di cui è stata data la definizione di colonna e rispetto alla lettura di una riga di lega 
            /// </summary>
            public static string WARNING_MANCATALETTURACOMPLETAPROPRIETACONCENTRAZIONIELEMENTO = "WARNING - riga {0}: le proprieta relative a un certo elemento sono state lasciate completamente vuote, il caso è corretto se l'elemento non fa parte della definizione per la lega corrente.\n";

            #endregion
        }
    }
}
