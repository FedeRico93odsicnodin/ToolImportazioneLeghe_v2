using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Utils
{
    /// <summary>
    /// Messaggi di errore che è possibile incontrare utilizzando il programma nelle 2 versioni
    /// </summary>
    public static class ExceptionMessages
    {
        #region LOGGING

        /// <summary>
        /// Segnalazione di non poter continuare per via di una mancata ricostruzione di parte del path per 
        /// la ricreazione del file di log 
        /// </summary>
        public static string LOGGING_NONHOTROVATOPARTEDELPATH = "impossibile ricreare il file di log, non è stata trovata parte del path '{0}'";


        /// <summary>
        /// Segnalazione di non aver trovato il file di log sul quale andare a scrivere una certa procedura 
        /// </summary>
        public static string LOGGING_INVALIDLOGPATH = "non ho trovato il log '{0}' file sul quale si è chiesto di scrivere un comando";


        /// <summary>
        /// Segnalazione generica di errore nel tentativo di configurazione della procedura di logging
        /// </summary>
        public static string LOGGING_VERIFICADIUNCERTOERRORE = "ho riscontrato il seguente errore durante la configurazione del LOGGING\n\n{0}";

        #endregion


        #region UTIL FUNCTIONS

        /// <summary>
        /// Segnalazione di un path nullo per l'ottenimento del nome proprio del file 
        /// </summary>
        public static string UTILFUNCTIONS_PATHNULL = "il path per un certo file è nullo o vuoto";


        /// <summary>
        /// Segnalazione di un nome nullo attribuito al file a un certo path che comunque è stato analizzato
        /// </summary>
        public static string UTILFUNCTIONS_NOMEVUOTO = "il nome attribuito al file è nullo";


        /// <summary>
        /// Segnalazione di errore nel tentativo di ricavare il nome per un determinato file 
        /// </summary>
        public static string UTILFUNCTIONS_ERRORENELLALETTURANOMEFILE = "ho riscontrato un errore nel tentativo di ricavare un nome per un file\n\n{0}";

        #endregion


        #region EXCEL 

        /// <summary>
        /// Eccezione relativa all'assenza di path per il file excel corrente 
        /// </summary>
        public static string EXCEL_EMPTYPATH = "il path passato per la lettura del file excel non corrisponde a nessuna informazione";


        /// <summary>
        /// Eccezione di inesistenza del file excel che viene passato in apertura per leggere le informazioni
        /// </summary>
        public static string EXCEL_SOURCENOTEXISTING = "il file excel passato come sorgente non esiste";


        /// <summary>
        /// Eccezione relativa alla non definizione per il formato relativo al file excel in apertura corrente 
        /// </summary>
        public static string EXCEL_FORMATNOTDEFINED = "il formato per il file excel in apertura corrente non è stato definito";


        /// <summary>
        /// Eccezione mancata apertura corretta per il foglio excel passato in input
        /// </summary>
        public static string EXCEL_PROBLEMAAPERTURAFILE = "si è verificato il seguente problema durante l'apertura del file excel '{0}':\n\n{1}";


        /// <summary>
        /// Eccezione relativa al fatto che il file excel corrente non è stato correttamente predisposto in memoria per l'analisi
        /// </summary>
        public static string EXCEL_FILENOTINMEMORY = "il file excel per l'analisi non è stato correttamente caricato in memoria";


        /// <summary>
        /// Eccezione utilizzata a livello di creazione e inserimento proprietà excel in lettura all'interno del wrapper
        /// </summary>
        public static string EXCEL_READINGPROPERTIES = "si è verificato un errore di implementazione durante la creazione istanza di proprieta excel";


        /// <summary>
        /// Ricerca per proprieta contenuta nel wrapper non andata a buon fine 
        /// </summary>
        public static string EXCEL_PROPERTYNOTDEFINED = "la proprieta ricercata durante la fase di analisi non è stata definita all'interno del wrapper per '{0}'";


        /// <summary>
        /// Eccezione relativa al fatto che manca l'header di colonna nelle definizioni obbligatorie e sulla quale andare a riconoscere la presenza di eventuali 
        /// elementi sottostanti
        /// </summary>
        public static string EXCEL_COLCRITERINONPRESENTE = "non posso proseguire l'analisi per individuare se il foglio corrente è di concentrazioni: manca la definizione della colonna su cui distinguere gli elementi (CRITERI)";

        #endregion
    }
}
