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
        /// Classe contenente le diverse intestazioni disponibili per le diverse tipologie di foglio excel e una analisi che non va a buon fine 
        /// per un determinato foglio in iterazione corrente 
        /// </summary>
        public static class Headers_ExcelSheet_ErrorMessages
        {

            /// <summary>
            /// Permette di ritornare l'intestazione da dare prima di inserire errori per un determinato foglio excel appartenenti a file del primo formato 
            /// relativamente alla validazione per il primo tipo (proprieta di lega)
            /// </summary>
            /// <param name="sheetName"></param>
            /// <returns></returns>
            public static string GetHeader_RecognizeSheet_Type1Format1(string sheetName)
            {
                return "=================================================================\n" +
                    String.Format("RICONOSCIMENTO PROPRIETA DI LEGHE - FOGLIO : '{0}'\n", sheetName) +
                    "=================================================================\n";
            }


            /// <summary>
            /// Permette di ritornare l'intestazione da dare prima di inserire errori per un determinato foglio excel appartenente al file del primo formato 
            /// relativamente alla validazione per il secondo tipo (proprieta di concentrazioni e quadranti di concentrazioni)
            /// </summary>
            /// <param name="sheetName"></param>
            /// <returns></returns>
            public static string GetHeader_RecognizeSheet_Type2Format1(string sheetName)
            {
                return "=================================================================\n" +
                    String.Format("RICONOSCIMENTO PROPRIETA DI CONCENTRAZIONI - FOGLIO : '{0}'\n", sheetName) +
                    "=================================================================\n";
            }

            
        }



        /// <summary>
        /// Messaggi di errore per la prima tipologia di foglio per il primo formato
        /// </summary>
        public static class Formato1_Foglio1_Leghe
        {

            #region RICONOSCIMENTO FOGLIO 

            /// <summary>
            /// Mancata lettura di un header per le proprieta di lega per il primo formato e il primo foglio disponibile
            /// </summary>
            public static string ERRORE_MANCATORICONOSCIMENTOPROPRIETAHEADERLEGHE = "ERRORE - non sono riuscito a trovare la proprieta '{0}' obbligatoria, relativamente al foglio '{1}' per la lettura delle proprieta di lega.\n";

            #endregion 


            #region RECUPERO - VALIDAZIONE 1 - INFORMAZIONI

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

            #endregion
        }


        /// <summary>
        /// Messaggi di errore per la seconda tipologia di foglio per il primo formato
        /// </summary>
        public static class Formato1_Foglio2_Concentrazioni
        {
            #region RICONOSCIMENTO FOGLIO 

            /// <summary>
            /// Errore di mancato riconoscimento del title per il quadrante delle concentrazioni in analisi corrente 
            /// </summary>
            public static string ERRORE_MANCATOTITLEQUADRANTE = "ERRORE - riga {0}: non sono riuscito a riconoscere un TITLE per l'eventuale quadrante da cui recuperare i valori delle concentrazioni";
            

            /// <summary>
            /// Errore di mancato riconscimento di una proprieta obbligatoria di header per le concentrazioni su un quadrante per il tipo di foglio in riconoscimento
            /// </summary>
            public static string ERRORE_MANCATORICONOSCIMENTOPROPRIETAHEADEROBBLIGATORIA = "ERRORE - riga {0}: non sono riuscito a riconoscere la proprieta {1} obbligatoria per le concentrazioni";


            /// <summary>
            /// Errore di mancato riconoscimento di almeno un elemento per il quadrante delle concentrazioni corrente 
            /// </summary>
            public static string ERRORE_NESSUNRICONOSCIMENTOPERELEMENTO = "ERRORE - non sono riuscito a riconoscere nessun elemento per eventuale completamento concentrazioni quadrante a partire dalla riga {0}";


            /// <summary>
            /// Errore di mancato riconoscimento di alcun quadrante per poter proseguire correttamente l'analisi e l'eventuale lettura delle concentrazioni per i diversi materiali contenuti in questo foglio
            /// </summary>
            public static string ERRORE_NESSUNQUADRANTECONCENTRAZIONIUTILEPERANALISI = "ERRORE - la tipologia per il foglio corrente è per la lettura delle concentrazioni, ma nessun quadrante di concentrazioni contiene delle informazioni utili per continuare l'analisi";

            #endregion
            

            #region RECUPERO - VALIDAZIONE 1 - INFORMAZIONI

            /// <summary>
            /// Errore di mancata lettura del nome per il quadrante delle concentrazioni correntemente in analisi per la situazione corrente
            /// </summary>
            public static string ERRORE_NOMEMATERIALELETTURAQUADRANTEVUOTO = "ERRORE - riga {0}, colonna {1}: nessuna informazione attribuita al nome per il MATERIALE per questo quadrante.\n";


            /// <summary>
            /// Errore durante l'analisi delle proprieta obbligatorie relative al quadrante di concentrazione corrente e per quanto riguarda le 
            /// proprieta relative alle diverse concentrazioni applicate per la lega corrente 
            /// </summary>
            public static string ERRORE_MANDATORYPROPERTYMANCANTE_CONCENTRAZIONI = "ERRORE - riga {0}: nessuna inforamzione valorizzata per la proprieta obbligatoria '{1}', le concentrazioni per la lega non verranno prese in considerazione per le analisi successive.\n";

            #endregion
        }


        /// <summary>
        /// Messaggi di errore per la prima tipologia di foglio e il secondo formato disponibile
        /// </summary>
        public static class Formato2_Foglio1_LegheConcentrazioni
        {
            #region RICONOSCIMENTO FORMATO 2 

            /// <summary>
            /// Errore durante l'analisi di riga e le proprieta direttamente collegate alla lega, in particolare c'è il mancato riconoscimento per una delle proprieta obbligatorie
            /// queste proprieta sono invalidanti rispetto alla riga e quindi alla lega corrente per le analisi e le valorizzazioni successive
            /// </summary>
            public static string ERRORE_MANCATORICONOSCIMENTOPROPRIETAOBBLIGATORIALEGA = "ERRORE - riga {0}: mancato riconoscimento per la proprieta OBBLIGATORIA DI LEGA '{1}'";


            /// <summary>
            /// Segnalazione di mancato riconoscimento per almeno una lega e per le sue proprieta globali e di concentrazione per il foglio corrente 
            /// </summary>
            public static string ERRORE_NESSUNAINFORMAZIONEDIRIGALETTA = "ERRORE - per il foglio '{0}' non ho letto ALCUNA INFORMAZIONE DI RIGA (proprieta leghe / concentrazioni)";


            /// <summary>
            /// Errore per il mancato riconoscimento di una proprieta obbligatoria per le concentrazioni, questa proprieta non riconosciuta è invalidante rispetto a tutto il processo successivo 
            /// di analisi per il foglio corrente 
            /// </summary>
            public static string ERRORE_MANCATORICONOSCIMENTOPROPRIETAOBBLIGATORIACONCENTRAZIONI = "ERRORE - riga {0}: mancato riconoscimento per la proprieta OBBLIGATORIA DI CONCENTRAZIONE '{1}'";


            /// <summary>
            /// Errore relativo al mancato riconoscimento delle colonne di header per il foglio di secondo formato attuale, questa opzione è invalidante rispetto a 
            /// tutta la restante iterazione possibile 
            /// </summary>
            public static string ERRORE_NESSUNACONCENTRAZIONERICONOSCIUTAPERHEADER = "ERRORE - per il foglio '{0}' non ho letto ALCUNA INFORMAZIONE DI HEADER COLONNE CONCENTRAZIONI";

            #endregion
            

            #region RECUPERO INFORMAZIONI - VALIDAZIONE 1

            /// <summary>
            /// Errore di mancata lettura per il nome di un elemento per l'insieme di colonne che mi caratterizzano le sue proprieta inerenti alle leghe in lettura corrente 
            /// </summary>
            public static string ERRORE_MANCATALETTURANOMEELEMENTO = "ERRORE - riga {0}, colonna {1}: nessuna informazione letta per il NOME dell'elemento sul quale andare a leggere le concentrazioni correnti.\n";


            /// <summary>
            /// Errore mancata lettura di una proprieta obbligatoria per la lettura delle proprieta obbligatorie relative ai parametri principali di lega
            /// </summary>
            public static string ERRORE_MANCATALETTURAPROPRIETAOBBLIGATORIA_LEGA = "ERRORE - riga {0}, colonna {1}: il valore per la proprieta '{2}' OBBLIGATORIA per le LEGHE è NULL.\n";


            /// <summary>
            /// Errore mancata lettura di una proprieta obbligatoria per la lettrua delle proprieta obbligatorie relative ai parametri principali di concentrazioni
            /// </summary>
            public static string ERRORE_MANCATALETTURAPROPRIETAOBBLIGATORIA_CONCENTRAZIONI = "ERRORE - riga {0}, colonna {1}: il valore per la proprieta '{2}' OBBLIGATORIA per le CONCENTRAZIONI è NULL.\n";


            /// <summary>
            /// Errore mancata lettura di concentrazioni per una delle leghe inserite nel foglio excel corrente di formato 2
            /// </summary>
            public static string ERRORE_NESSUNACOLONNACONCENTRAZIONIPERLEGACORRENTE = "ERRORE - riga {0}: non ho trovato nessun valore da leggere per le concentrazioni e per le informazioni relative ad una delle leghe inserite";
            

            /// <summary>
            /// Errore di non poter continuare con analisi del foglio corrente in quanto mancano tutte le informazioni di base per poter leggere almeno una lega contenuta nel foglio 
            /// </summary>
            public static string ERRORE_ANALISINTERROTTAPERTUTTEPROPRIETALEGHEMANCANTI = "ERRORE - non posso continuare con l'analisi del foglio excel '{0}' perché mancano tutte le informazioni di LEGA per poter proseguire.\n";


            /// <summary>
            /// Errore di non poter continuare con analisi del foglio excel in quanto mancano tutte le informazioni di base per i valori in lettura per le concentrazioni per poter proseguire 
            /// </summary>
            public static string ERRORE_ANALISIINTERROTTAPERTUTTECONCENTRAZIONIMANCANTI = "ERRORE - non posso continuare con l'analisi del foglio excel '{0}' perché mancano tutte le informazioni per le CONCENTRAZIONI DA LEGGERE per poter proseguire.\n";


            /// <summary>
            /// Erroe di mancata lettura per alcune delle leghe di informazioni di carattere generale di lega e per alcune delle leghe delle informazioni su tutti gli elementi per proseguire 
            /// in ogni caso questi 2 insiemi formano uno unico che non permette di proseguire con l'analisi per nessuno degli elementi
            /// </summary>
            public static string ERRORE_PERTUTTELELEGHESITUAZIONEMISTANONLETTURAPROPRIETALEGHECONCENTRAZIONI = "ERRORE - non posso continuare con l'analisi del foglio excel '{0}' perché per alcune leghe manca la lettura delle informazioni generali e per altre la lettura delle concentrazioni fondamentali per proseguire.\n";

            #endregion
        }
    }
}
